# -*- coding: utf-8 -*-
"""
대시보드 데이터 연동 및 정적 갱신 스크립트 (update_index.py)

[의도]
1. 사용자가 선택한 '정적 빌드 방식'에 따라 Supabase DB의 최신 5월 데이터를 집계하여 index.html에 정적으로 반영합니다.
2. Supabase DB가 무료 티어 사양으로 인해 쿼리 제한이나 타임아웃이 발생할 수 있는 리스크가 있습니다.
3. 이를 예방하기 위해, Supabase API를 통한 전체 로깅 데이터 로드가 실패하면 로컬에 존재하는 '08_5월.xlsx' 파일로부터 직접 데이터를 읽어 집계하는 "하이브리드 예외 처리(Fallback)" 구조를 구현합니다.
4. 집계 결과를 바탕으로 index.html의 요약 KPI 카드, 상위 10인 사용 현황 바, 조직별 데이터 맵핑 변수(var OD)를 정밀하게 교체합니다.
"""

import os
import re
import json
import openpyxl
from datetime import datetime
from collections import Counter, defaultdict
from supabase import create_client, Client

# Supabase 연결 설정
URL = "https://seanzwnadqaneusqeami.supabase.co"
KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNlYW56d25hZHFhbmV1c3FlYW1pIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEzNzMxNjIsImV4cCI6MjA5Njk0OTE2Mn0.3-R97YJzSsVW2ecJSW5briUFwFVNAATJHhASB3xgNuI"

def fetch_data_from_supabase():
    """Supabase에서 데이터를 페이지네이션을 통해 1000행씩 안전하게 조회하여 메모리에 병합합니다."""
    print("[1] Supabase로부터 실시간 데이터 로딩 시도...")
    supabase: Client = create_client(URL, KEY)
    
    all_data = []
    limit = 1000
    offset = 0
    
    # 필요한 컬럼만 명시적으로 선택하여 페이로드 크기를 대폭 축소하고 전송 속도를 높입니다.
    cols = "userNo,userName,chatDate,chatHour,team_region,userAction"
    
    while True:
        try:
            # 페이지네이션 시 반드시 id 기준으로 정렬하여 행 중복/누락을 방지합니다.
            res = supabase.table("chat_logs").select(cols).order("id").range(offset, offset + limit - 1).execute()
            if not res.data:
                break
            all_data.extend(res.data)
            if len(res.data) < limit:
                break
            offset += limit
            if offset % 10000 == 0:
                print(f"  -> {offset} 행 다운로드 완료...")
        except Exception as e:
            print(f"[Warning] Supabase 데이터 로드 중 오류 발생 (Pagination 중단): {e}")
            return None
            
    print(f" -> Supabase 로드 완료: 총 {len(all_data)} 행")
    return all_data

def fetch_data_from_excel(file_path="08_5월.xlsx"):
    """Supabase에 장애가 있거나 타임아웃 발생 시 로컬 엑셀 파일에서 데이터를 스트리밍하여 직접 집계합니다."""
    print(f"[1] 로컬 엑셀 파일로부터 데이터 로딩 시도 (Fallback): {file_path}")
    if not os.path.exists(file_path):
        print(f"[Error] 로컬 백업용 엑셀 파일이 없습니다: {file_path}")
        return None
        
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active
    
    row_iter = ws.iter_rows(values_only=True)
    headers = next(row_iter)
    col_map = {col_name: i for i, col_name in enumerate(headers)}
    
    all_data = []
    for row in row_iter:
        if not row or (row[col_map["userNo"]] is None and row[col_map["userAction"]] is None):
            continue
            
        # Helper to safely retrieve values
        def get_val(primary, fallback=None, default=""):
            if primary in col_map:
                val = row[col_map[primary]]
                return str(val).strip() if val is not None else default
            if fallback and fallback in col_map:
                val = row[col_map[fallback]]
                return str(val).strip() if val is not None else default
            return default

        chat_date = ""
        if "chatDate" in col_map:
            val = row[col_map["chatDate"]]
            chat_date = str(val)[:10] if val is not None else ""
        elif "chatTime" in col_map:
            val = row[col_map["chatTime"]]
            chat_date = str(val)[:10] if val is not None else ""

        chat_hour = 0
        if "chatHour" in col_map:
            val = row[col_map["chatHour"]]
            try:
                chat_hour = int(val) if val is not None else 0
            except:
                chat_hour = 0
        elif "chatTime" in col_map:
            val = row[col_map["chatTime"]]
            if val:
                try:
                    time_str = str(val).strip().split()[-1]
                    chat_hour = int(time_str.split(":")[0])
                except:
                    chat_hour = 0
                    
        all_data.append({
            "userNo": get_val("userNo"),
            "userName": get_val("userName"),
            "chatDate": chat_date,
            "chatHour": chat_hour,
            "team_region": get_val("team_region"),
            "userAction": get_val("userAction")
        })
    print(f" -> 로컬 엑셀 로드 완료: 총 {len(all_data)} 행")
    return all_data

def process_and_update():
    # 1. 데이터 소스 획득 (하이브리드 조회 방식)
    data = fetch_data_from_supabase()
    if not data or len(data) == 0:
        print("[Warning] Supabase 데이터가 유효하지 않아 로컬 엑셀 파일 기반으로 데이터 집계를 수행합니다.")
        data = fetch_data_from_excel()
        
    if not data:
        print("[Fatal Error] 데이터를 수집할 수 없어 대시보드 갱신을 중단합니다.")
        return

    print("[2] 비즈니스 로직 기반 데이터 통계 분석 시작...")
    
    total_events = len(data)
    
    # 통계 계산을 위한 임시 자료구조
    unique_users = set()
    total_questions = 0
    total_clicks = 0
    
    user_event_counter = Counter()
    user_q_counter = Counter()
    user_name_map = {}
    
    # 부서별 집계용 임시 딕셔너리
    # { dept_name: { 'users': set(), 'total': 0, 'q': 0, 'click': 0 } }
    dept_stats = defaultdict(lambda: {"users": set(), "total": 0, "q": 0, "click": 0})
    
    # 활성 사용자 (검색을 1회 이상 수행한 사용자) 구별용
    active_users = set()
    
    # 109개 센터 계산용
    unique_centers = set()
    
    for row in data:
        user_no = row.get("userNo") or ""
        user_name = row.get("userName") or ""
        dept_name = row.get("team_region") or ""
        action = row.get("userAction") or ""
        
        # 사번이 공백이거나 유효하지 않은 값 무시 방지
        if not user_no:
            continue
            
        unique_users.add(user_no)
        user_name_map[user_no] = user_name
        user_event_counter[user_no] += 1
        
        if dept_name:
            # 부서 세부 매핑
            dept_stats[dept_name]["total"] += 1
            dept_stats[dept_name]["users"].add(user_no)
            
        if action == "QUESTION":
            total_questions += 1
            user_q_counter[user_no] += 1
            active_users.add(user_no)
            if dept_name:
                dept_stats[dept_name]["q"] += 1
                
        elif action == "LINK_CLICK":
            total_clicks += 1
            if dept_name:
                dept_stats[dept_name]["click"] += 1
                
    # 1인당 평균 사용량
    avg_usage = round(total_events / len(unique_users), 1) if unique_users else 0
    # 활성 사용자 비율
    active_rate = round(len(active_users) / len(unique_users) * 100, 1) if unique_users else 0
    # 검색->클릭 전환율
    click_rate = round(total_clicks / total_questions * 100, 1) if total_questions else 0
    
    # 상위 10명 사용자 추출 및 비율 계산
    top_10_users = user_event_counter.most_common(10)
    top_10_sum = sum(cnt for _, cnt in top_10_users)
    top_10_share = round(top_10_sum / total_events * 100, 1) if total_events else 0
    
    print(f" - 전체 사용자 수 (Unique): {len(unique_users)} 명")
    print(f" - 총 이벤트 수: {total_events} 건")
    print(f" - 총 검색 수 (QUESTION): {total_questions} 건")
    print(f" - 링크 클릭 수: {total_clicks} 건")
    print(f" - 활성 사용자 수: {len(active_users)} 명 ({active_rate}%)")
    print(f" - 검색->클릭 전환율: {click_rate}%")
    print(f" - 상위 10인 점유율: {top_10_share}%")

    # 조직 Roster (이전 대시보드에 기재된 조직 인원 로스터 기준)
    rosters = {
        "고객가치혁신한국수도권중부담당": 1238,
        "고객가치혁신한국서남부담당": 982,
        "고객가치혁신한국충청실": 270,
        "고객가치혁신한국강원실": 105,
        "(미지정)": 353
    }
    
    # 부서(조직) 매핑 JSON 데이터 구성 (var OD)
    od_data = {}
    for dept, stats in dept_stats.items():
        # 대시보드와 정확하게 일치하는 대분류 조직만 필터링 또는 매핑
        clean_dept = dept.strip()
        if not clean_dept:
            clean_dept = "(미지정)"
            
        users_cnt = len(stats["users"])
        total_cnt = stats["total"]
        q_cnt = stats["q"]
        click_cnt = stats["click"]
        
        per_user = round(total_cnt / users_cnt, 1) if users_cnt else 0
        per_user_q = round(q_cnt / users_cnt, 1) if users_cnt else 0
        
        roster_size = rosters.get(clean_dept, users_cnt) # 로스터 정보 없으면 현재 참여자로 매핑
        part_rate = round(users_cnt / roster_size * 100, 1) if roster_size else 100.0
        if part_rate > 100.0:
            part_rate = 100.0 # 100%를 초과할 수 없음
            
        od_data[clean_dept] = {
            "team": clean_dept,
            "users": users_cnt,
            "total": total_cnt,
            "q": q_cnt,
            "click": click_cnt,
            "perUser": per_user,
            "perUserQ": per_user_q,
            "partRate": part_rate
        }
    
    # (미지정) 및 빠진 조직에 대한 기본 세팅 보강
    for std_dept in rosters.keys():
        if std_dept not in od_data:
            od_data[std_dept] = {
                "team": std_dept,
                "users": 0,
                "total": 0,
                "q": 0,
                "click": 0,
                "perUser": 0.0,
                "perUserQ": 0.0,
                "partRate": 0.0
            }

    print("[3] HTML 갱신 처리 작업 시작...")
    html_file = "index.html"
    with open(html_file, "r", encoding="utf-8") as f:
        html = f.read()
        
    # 3-1. 조직 필터 데이터 교체 (var OD = { ... })
    od_pattern = r"var OD = \{.*?\};"
    new_od_str = f"var OD = {json.dumps(od_data, ensure_ascii=False)};"
    html = re.sub(od_pattern, new_od_str, html)
    print(" -> var OD 데이터 교체 완료.")
    
    # 3-2. 헤더 메타 칩 업데이트
    # <span class="chip-badge">🗂️ 173,034 이벤트</span>
    html = re.sub(r'🗂️\s*[\d,]+\s*이벤트', f"🗂️ {total_events:,} 이벤트", html)
    # <span class="chip-badge">👥 2,740 사용자</span>
    html = re.sub(r'👥\s*[\d,]+\s*사용자', f"👥 {len(unique_users):,} 사용자", html)
    print(" -> 헤더 메타 뱃지 업데이트 완료.")

    # 3-3. KPI 요약 카드 부분 교체
    # KPI 1: 전체 사용자 (Unique)
    # <div class="kpi-v">2,740</div>
    kpi_user_pattern = r'(<div class="kpi"[^>]*?>.*?<div class="kpi-ic">👥</div>.*?<div class="kpi-v">)([\d,]+)(</div>)'
    html = re.sub(kpi_user_pattern, rf"\g<1>{len(unique_users):,}\g<3>", html, flags=re.DOTALL)
    
    # KPI 2: 총 이벤트 수 (Primary KPI)
    # <div class="kpi primary-kpi">...⚡...<div class="kpi-v">173,034</div>
    kpi_event_pattern = r'(<div class="kpi primary-kpi">.*?⚡.*?<div class="kpi-v">)([\d,]+)(</div>)'
    html = re.sub(kpi_event_pattern, rf"\g<1>{total_events:,}\g<3>", html, flags=re.DOTALL)
    
    # KPI 3: 총 검색 수 (QUESTION)
    # <div class="kpi-v">52,270</div>
    kpi_q_pattern = r'(<div class="kpi"[^>]*?>.*?🔎.*?<div class="kpi-v">)([\d,]+)(</div>)'
    html = re.sub(kpi_q_pattern, rf"\g<1>{total_questions:,}\g<3>", html, flags=re.DOTALL)
    
    # KPI 4: 링크 클릭 수
    # <div class="kpi-v">18,633</div>
    kpi_click_pattern = r'(<div class="kpi"[^>]*?>.*?🖱️.*?<div class="kpi-v">)([\d,]+)(</div>)'
    html = re.sub(kpi_click_pattern, rf"\g<1>{total_clicks:,}\g<3>", html, flags=re.DOTALL)
    
    # KPI 5: 1인당 평균 사용량
    # <div class="kpi-v">63.2</div>
    kpi_avg_pattern = r'(<div class="kpi"[^>]*?>.*?📊.*?<div class="kpi-v">)([\d\.,]+)(</div>)'
    html = re.sub(kpi_avg_pattern, rf"\g<1>{avg_usage}\g<3>", html, flags=re.DOTALL)
    
    # KPI 6: 활성 사용자 비율
    # <div class="kpi primary-kpi">...✅...<div class="kpi-v">93.5%</div>
    kpi_act_pattern = r'(<div class="kpi primary-kpi">.*?✅.*?<div class="kpi-v">)([\d\.,%]+)(</div>)'
    html = re.sub(kpi_act_pattern, rf"\g<1>{active_rate}%\g<3>", html, flags=re.DOTALL)
    
    # KPI 7: 검색->클릭 전환율
    # <div class="kpi-v">35.6%</div>
    kpi_conv_pattern = r'(<div class="kpi"[^>]*?>.*?🔄.*?<div class="kpi-v">)([\d\.,%]+)(</div>)'
    html = re.sub(kpi_conv_pattern, rf"\g<1>{click_rate}%\g<3>", html, flags=re.DOTALL)
    
    # KPI 8: 상위 10명 점유율
    # <div class="kpi-v">5.5%</div>
    kpi_top_pattern = r'(<div class="kpi"[^>]*?>.*?🏆.*?<div class="kpi-v">)([\d\.,%]+)(</div>)'
    html = re.sub(kpi_top_pattern, rf"\g<1>{top_10_share}%\g<3>", html, flags=re.DOTALL)
    
    print(" -> KPI 요약 수치 카드 업데이트 완료.")

    # 3-4. 상위 사용자 Top 10 바 리스트 업데이트
    # <div class="hbars"> 내부 데이터 구성
    hbars_html = ""
    max_val = top_10_users[0][1] if top_10_users else 1
    
    for user_no, total_cnt in top_10_users:
        user_name = user_name_map.get(user_no, "미식별")
        q_cnt = user_q_counter.get(user_no, 0)
        percentage = round((total_cnt / max_val) * 100, 1)
        
        hbars_html += f"""                        <div class="hb-row">
                            <div class="hb-label" title="{user_name}">{user_name}</div>
                            <div class="hb-track">
                                <div class="hb-fill" style="width:{percentage}%;background:var(--brand)"></div><span
                                    class="hb-val">{total_cnt:,}<span class="hb-sub"> (검색 {q_cnt:,})</span></span>
                            </div>
                        </div>\n"""
                        
    # HTML 내 <!-- TOP_10_USERS_START --> ... <!-- TOP_10_USERS_END --> 영역 교체
    hbars_regex = r'(<!-- TOP_10_USERS_START -->)(.*?)(<!--\s*TOP_10_USERS_END\s*-->)'
    html = re.sub(hbars_regex, rf"\g<1>\n{hbars_html}                        \g<3>", html, flags=re.DOTALL)
    print(" -> Top 10 사용자 목록 차트 업데이트 완료.")

    # 3-5. 분석 기간 날짜 세팅 및 변경 사항 저장
    # <div class="date">2026-05-01 ~ 2026-05-31</div>
    html = re.sub(r'<div class="date">.*?</div>', '<div class="date">2026-05-01 ~ 2026-05-31</div>', html)
    
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html)
        
    print(f"\n[성공] 대시보드 파일({html_file})이 5월 분석 데이터 기준으로 완벽하게 업데이트되었습니다!")

if __name__ == "__main__":
    process_and_update()
