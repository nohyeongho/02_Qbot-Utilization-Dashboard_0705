# -*- coding: utf-8 -*-
"""
Supabase 데이터 적재 스크립트 (ingest_supabase.py)

[의도]
1. 17만 건이 넘는 대용량 엑셀 파일(08_5월.xlsx)을 판다스로 한 번에 읽으면 MemoryError(OOM)가 발생할 수 있습니다.
2. 이를 방지하기 위해 openpyxl의 read_only=True 모드를 활용해 메모리 사용량을 최소화하며 데이터를 스트리밍 방식으로 한 행씩 읽어옵니다.
3. 데이터를 5,000행 단위의 청크(Chunk)로 나누어 Supabase의 chat_logs 테이블에 벌크 인서트(Bulk Insert)합니다.
4. 인서트 전 기존 데이터를 안전하게 삭제하여 데이터 중복을 방지합니다.
"""

import os
import time
from datetime import datetime
import openpyxl
from supabase import create_client, Client

# Supabase 연결 설정 정보
URL = "https://seanzwnadqaneusqeami.supabase.co"
KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNlYW56d25hZHFhbmV1c3FlYW1pIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEzNzMxNjIsImV4cCI6MjA5Njk0OTE2Mn0.3-R97YJzSsVW2ecJSW5briUFwFVNAATJHhASB3xgNuI"

# Supabase 클라이언트 초기화
supabase: Client = create_client(URL, KEY)

def clean_value(val):
    """None 값을 빈 문자열로 정제하고 문자열 타입으로 일관되게 변환합니다."""
    if val is None:
        return ""
    return str(val).strip()

def clean_int(val):
    """정수형 데이터를 변환하며, 에러 발생 시 0을 반환합니다."""
    if val is None:
        return 0
    try:
        return int(val)
    except:
        return 0

def format_date(val):
    """날짜 필드를 YYYY-MM-DD 형식의 문자열로 변환합니다."""
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    # 문자열에서 날짜 형태만 추출
    val_str = str(val).strip()
    if len(val_str) >= 10:
        return val_str[:10]
    return val_str

def ingest_data(file_path: str):
    if not os.path.exists(file_path):
        print(f"[Error] 파일이 존재하지 않습니다: {file_path}")
        return

    print(f"[Step 1] 엑셀 파일 로딩 시작 (Read-Only 모드): {file_path}")
    start_time = time.time()
    
    # read_only=True로 지정하여 대용량 파일을 메모리에 모두 올리지 않고 필요할 때만 디스크에서 읽어옵니다.
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active # 첫 번째 시트 (rawData) 활성화
    print(f"활성화된 시트명: {ws.title}")
    
    # 첫 행(Header)을 읽어 각 컬럼의 인덱스를 맵핑합니다.
    row_iter = ws.iter_rows(values_only=True)
    headers = next(row_iter)
    
    # 필요한 컬럼들의 위치 파악 (0-indexed)
    col_map = {col_name: i for i, col_name in enumerate(headers)}
    print("엑셀 컬럼 맵핑 완료.")
    
    # 기존 Supabase 데이터 삭제 (초기화)
    # [의도] 5월 데이터를 깨끗하게 다시 올리기 위해 기존에 적재되어 있던 모든 로그를 지웁니다.
    # 대량 삭제 시 포스트그레스 타임아웃(57014)을 피하기 위해 5,000건씩 끊어서 안전하게 지웁니다.
    try:
        print("[Step 2] Supabase 기존 데이터 초기화 중 (청크 단위 삭제)...")
        deleted_count = 0
        while True:
            # 5000개씩 ID를 가져와서 삭제
            res = supabase.table("chat_logs").select("id").limit(5000).execute()
            if not res.data:
                break
            ids_to_delete = [r["id"] for r in res.data]
            supabase.table("chat_logs").delete().in_("id", ids_to_delete).execute()
            deleted_count += len(ids_to_delete)
            print(f"  -> 기존 데이터 {deleted_count}개 삭제 완료...")
        print("기존 데이터 전체 삭제 완료.")
    except Exception as e:
        print(f"[Warning] 기존 데이터 삭제 중 에러 발생: {e}")

    # 데이터 추출 및 전송 루프
    records = []
    chunk_size = 3000 # 한 번에 Supabase API로 전송할 행 수 (네트워크 타임아웃 방지를 위해 3000행 지정)
    total_count = 0
    
    print("[Step 3] 데이터 스트리밍 및 적재 시작...")
    
    for row in row_iter:
        # 빈 행 패스 (사번과 행동 유형이 둘 다 없는 경우만 진짜 빈 행으로 취급하여 패스)
        if not row or (row[col_map["userNo"]] is None and row[col_map["userAction"]] is None):
            continue
            
        try:
            # 동적 컬럼 맵핑
            def get_col_val(primary_col, fallback_col=None, default=""):
                if primary_col in col_map:
                    val = row[col_map[primary_col]]
                    return clean_value(val)
                if fallback_col and fallback_col in col_map:
                    val = row[col_map[fallback_col]]
                    return clean_value(val)
                return default

            # 날짜 및 시간 특별 처리
            chat_date = ""
            if "chatDate" in col_map:
                chat_date = format_date(row[col_map["chatDate"]])
            elif "chatTime" in col_map:
                val = row[col_map["chatTime"]]
                if val:
                    chat_date = str(val)[:10]

            chat_hour = 0
            if "chatHour" in col_map:
                chat_hour = clean_int(row[col_map["chatHour"]])
            elif "chatTime" in col_map:
                val = row[col_map["chatTime"]]
                if val:
                    try:
                        # chatTime이 "14:30:22" 또는 "2026-05-01 14:30:22" 형식일 수 있음
                        time_str = str(val).strip().split()[-1]
                        chat_hour = int(time_str.split(":")[0])
                    except:
                        chat_hour = 0

            # chatTime 포맷팅 처리
            chat_time_val = None
            if "chatTime" in col_map:
                val = row[col_map["chatTime"]]
                if val:
                    if isinstance(val, datetime):
                        chat_time_val = val.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        chat_time_val = str(val).strip()

            # 엑셀의 값을 데이터베이스 스키마 구조에 맞추어 변환 및 정제 (새로운 엑셀 규격 스키마)
            record = {
                "chatMsg": get_col_val("chatMsg"),
                "chatTime": chat_time_val,
                "chatDate": chat_date,
                "chatHour": chat_hour,
                "userNo": get_col_val("userNo"),
                "userName": get_col_val("userName"),
                "userAction": get_col_val("userAction"),
                "userActionValue": get_col_val("userActionValue"),
                "questionTypeCd": get_col_val("questionTypeCd", default="기본"),
                "prodLvl2Cd": get_col_val("prodLvl2Cd"),
                "prodLvl2Name": get_col_val("prodLvl2Name"),
                "prodLvl3Cd": get_col_val("prodLvl3Cd"),
                "prodLvl3Name": get_col_val("prodLvl3Name"),
                "dept": get_col_val("dept"),
                "team_region": get_col_val("team_region"),
                "center": get_col_val("center"),
                "job_type": get_col_val("job_type")
            }
            records.append(record)
        except Exception as ex:
            print(f"[Warning] 행 파싱 중 예외 발생 (스킵): {ex}")
            continue

        # 청크 단위로 Supabase 적재 수행
        if len(records) >= chunk_size:
            try:
                supabase.table("chat_logs").insert(records).execute()
                total_count += len(records)
                print(f"적재 완료: {total_count} 행...")
            except Exception as e:
                print(f"[Error] Supabase 전송 실패: {e}")
                # 실패한 경우 재시도 처리 로직을 넣어 안전성을 높입니다.
                time.sleep(2)
                try:
                    supabase.table("chat_logs").insert(records).execute()
                    total_count += len(records)
                    print(f"재시도 적재 완료: {total_count} 행...")
                except Exception as re:
                    print(f"[Fatal] 재시도 실패, 스크립트 중단: {re}")
                    return
            records = [] # 청크 리스트 초기화

    # 남은 잔여 데이터 적재
    if records:
        try:
            supabase.table("chat_logs").insert(records).execute()
            total_count += len(records)
            print(f"적재 완료 (최종): {total_count} 행...")
        except Exception as e:
            print(f"[Error] 최종 잔여 데이터 적재 실패: {e}")

    end_time = time.time()
    elapsed = end_time - start_time
    print(f"\n[성공] 총 {total_count}건의 5월 로그 데이터가 Supabase에 적재되었습니다.")
    print(f"소요 시간: {round(elapsed, 1)}초")

if __name__ == "__main__":
    target_excel = "08_5월.xlsx"
    ingest_data(target_excel)
