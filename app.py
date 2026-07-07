import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import math

# 1. 페이지 설정
st.set_page_config(page_title="엘리베이터 원가산출", layout="wide")
st.title("🛗 엘리베이터 원가산출 프로그램")

# 2. 엑셀 데이터 로드
@st.cache_data(ttl=2)
def load_data():
    file_candidates = [
        "Elevator_Master_DB.csv", 
        "Elevator_Master_DB.xlsx",
        "Elevator_Expanded_Master_DB.xlsx - 통합단가마스터_확장형.csv"
    ]
    target_file = None
    for f in file_candidates:
        if os.path.exists(f):
            target_file = f
            break
            
    if not target_file:
        st.error("📁 폴더에 단가표 파일이 없습니다. 파일명을 확인해 주세요!")
        return pd.DataFrame()
        
    try:
        if target_file.endswith(".xlsx"):
            df_raw = pd.read_excel(target_file, header=None)
        else:
            try: df_raw = pd.read_csv(target_file, header=None, encoding="cp949")
            except Exception:
                try: df_raw = pd.read_csv(target_file, header=None, encoding="utf-8")
                except Exception: df_raw = pd.read_excel(target_file, header=None, engine="openpyxl")
        
        header_idx = 0
        for i, row in df_raw.iterrows():
            row_str = [str(x) for x in row.values]
            if any("구분" in s or "품목명" in s for s in row_str):
                header_idx = i
                break
        
        df = df_raw.iloc[header_idx+1:].copy()
        df.columns = [str(c).replace('\n', ' ').split('(')[0].strip() for c in df_raw.iloc[header_idx]]
        
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
                
        return df
    except Exception as e:
        st.error(f"❌ 파일을 읽는 중에 오류가 발생했습니다: {e}")
        return pd.DataFrame()

df_master = load_data()

def parse_cell_to_digits_list(text):
    if pd.isna(text): return []
    val_str = str(text).strip()
    if val_str in ["", "nan", "None"]: return []
    
    parts = re.split(r'[,/|]', val_str)
    digits_list = []
    for p in parts:
        dig = re.sub(r'[^0-9]', '', p)
        if dig: digits_list.append(dig)
    return digits_list

# OPB 등 비고란의 복잡한 층수 제약조건(예: 6,7,8,9 또는 10층 이상)을 정확히 매칭하는 헬퍼 함수
def match_floor_range_from_vigo(vigo_text, current_stop):
    vigo_clean = str(vigo_text).replace(" ", "")
    if not vigo_clean or vigo_clean in ["nan", "None"]:
        return True
    
    if "10층이상" in vigo_clean or "10이상" in vigo_clean:
        return current_stop >= 10
        
    base_part = vigo_clean.split("(")[0]
    digits = [int(d) for d in re.findall(r'\d+', base_part)]
    if digits:
        return current_stop in digits
        
    return True

if not df_master.empty:
    # 3. 사양 선택 UI
    st.markdown("### 📐 엘리베이터 사양 선택")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        floor_input = st.number_input("건물 층수 (FLOOR)", min_value=2, max_value=100, value=7)
        spec_구동 = st.selectbox("구동방식", [str(x).strip() for x in df_master["구동방식"].dropna().unique() if str(x).strip() not in ["", "nan"]])
        
        # [KeyError 대응 패치] 컬럼명이 판넬사양 / 판ネル사양 어떤 형태든 유연하게 인식하도록 수정
        target_panel_col = None
        for col in ["판넬사양", "판ネル사양", "판넬"]:
            if col in df_master.columns:
                target_panel_col = col
                break
        panel_opts = [str(x).strip() for x in df_master[target_panel_col].dropna().unique() if str(x).strip() not in ["", "nan"]] if target_panel_col else []
        spec_판넬 = st.selectbox("판넬사양", panel_opts)

    with col2:
        stop_input = st.number_input("서비스 STOP 수", min_value=2, max_value=100, value=7)
        
        inc_set = set()
        if "인승" in df_master.columns:
            for x in df_master["인승"].dropna().unique():
                for dig in parse_cell_to_digits_list(x):
                    inc_set.add(int(dig))
        all_inc_list = sorted(list(inc_set))
        spec_인승 = st.selectbox("인승 (숫자만)", [str(s) for s in all_inc_list], index=all_inc_list.index(13) if 13 in all_inc_list else 0)
        
        guide_opts = [str(x).strip() for x in df_master["가이드"].dropna().unique() if str(x).strip() not in ["", "nan"]] if "가이드" in df_master.columns else []
        spec_가이드 = st.selectbox("가이드 사양", guide_opts)

    with col3:
        sub_col1, sub_col2 = st.columns(2)
        with sub_col1:
            type_input = st.radio("출입구 형태", ["일방", "관통"], index=1, horizontal=True)
        with sub_col2:
            fire_input = st.radio("도어 방화 사양", ["일반도어", "방화도어"], index=1, horizontal=True)
        
        speed_set = set()
        if "속도" in df_master.columns:
            for x in df_master["속도"].dropna().unique():
                for dig in parse_cell_to_digits_list(x):
                    speed_set.add(int(dig))
        all_speed_list = sorted(list(speed_set))
        spec_속도 = st.selectbox("속도 (m/min)", [str(s) for s in all_speed_list], index=all_speed_list.index(60) if 60 in all_speed_list else 0)
        
        door_set = set()
        if "도어오프닝" in df_master.columns:
            for x in df_master["도어오프닝"].dropna().unique():
                for dig in parse_cell_to_digits_list(x):
                    door_set.add(int(dig))
        all_door_list = sorted(list(door_set))
        spec_오프닝 = st.selectbox("도어오프닝 사이즈 (mm)", [str(s) for s in all_door_list], index=all_door_list.index(900) if 900 in all_door_list else 0)

    st.markdown("---")

    # 4. 만능 수식 해독기
    def parse_qty_from_note(item_name, calc_type, note, base_stop, base_floor, gate_type, current_drive, spec_인승_선택값, is_roller_mode=False, is_fire_door=False):
        note_clean = str(note).replace(" ", "").upper() if pd.notna(note) else ""
        calc_clean = str(calc_type).replace(" ", "")
        item_clean = str(item_name).replace(" ", "")
        
        # [신규 예외 처리] 케이블_외 품목은 무조건 기본 1식(수량 1.0) 고정
        if "케이블_외" in item_clean or "케이블외" in item_clean:
            return 1.0

        if "MAIN_OPB" in item_clean:
            return 1.0

        # [로직 변경] CPI 수량: 관통이면 2개, 일방이면 1개 고정
        if "CPI" in item_clean:
            return 2.0 if gate_type == "관통" else 1.0

        if "SUB_OPB" in item_clean:
            if gate_type == "일방": 
                return 0.0
            return 1.0

        if "장애인용" in item_clean and "OPB" in item_clean:
            return 1.0
            
        if "조속기" in item_clean and "와이어" in item_clean:
            gov_wire_len = ((base_floor * 3) + 15) * 2
            return float(gov_wire_len)
            
        if "와이어" in item_clean and "소켓" in item_clean:
            insung_val = int(spec_인승_선택값)
            if current_drive.upper() == "MR":
                if insung_val == 8: return 6.0
                elif insung_val == 10: return 8.0
                else: return 10.0
            else:
                if insung_val <= 13: return 10.0
                elif insung_val <= 21: return 12.0
                else: return 14.0
            
        if "와이어" in item_clean and "소켓" not in item_clean and "조속기" not in item_clean:
            single_wire_len = (base_floor * 3) + 9
            insung_val = int(spec_인승_선택값)
            if insung_val <= 12: cols = 4
            elif insung_val == 13: cols = 5
            else: cols = 6
            return float(single_wire_len * cols)
        
        if "가이드레일" in item_clean:
            rail_qty = math.ceil((((base_floor * 3000) + 6000) / 5000) * 2)
            return float(rail_qty)
            
        if "레일조인트" in item_clean or "조인트" in item_clean:
            rail_qty = math.ceil((((base_floor * 3000) + 6000) / 5000) * 2)
            return float(max(0, rail_qty - 1))

        if "방화도어" in item_clean:
            if not is_fire_door: return 0.0
        
        if "가이드슈" in item_clean:
            return 4.0
            
        if "오일통" in item_clean or "오일받이" in item_clean:
            if is_roller_mode: return 0.0
            return 1.0
            
        if "구동방식" in calc_clean:
            match = re.search(rf'{current_drive.upper()}(\d+)', note_clean)
            if match: return float(match.group(1))
            return 1.0
        elif "STOP" in calc_clean:
            match = re.search(r'STOP([+-]\d+)', note_clean)
            if match: return max(0.0, float(base_stop) + float(match.group(1)))
            return float(base_stop)
        elif "층수" in calc_clean or "가이드" in calc_clean:
            match = re.search(r'FLOOR([+-]\d+)', note_clean)
            if match: return max(0.0, float(base_floor) + float(match.group(1)))
            return float(base_floor)
        elif "관통" in calc_clean:
            match_gate = re.search(rf'{gate_type.upper()}(\d+)', note_clean)
            if match_gate: return float(match_gate.group(1))
            return 2.0 if gate_type == "관통" else 1.0
        else:
            return 1.0

    # 5. 원가 계산 실행 버튼
    if st.button("📊 실시간 원가 계산 실행", type="primary", use_container_width=True):
        raw_bom_rows = []
        all_unique_items = [x for x in df_master["품목명"].dropna().unique() if str(x).strip() not in ["", "nan", "품목명"]]
        
        for item in all_unique_items:
            df_item = df_master[df_master["품목명"] == item]
            
            matched_row = None
            is_perfect_match = False
            best_fallback_row = None
            min_score = float('inf')
            status_text = "정상"
            # 💡 [여기서부터 수정] CPI 예외 로직 추가
            is_cpi_item = "CPI" in str(item).replace(" ", "").upper()
            
            for _, row in df_item.iterrows():
                # CPI는 다른 사양 검사 없이 즉시 통과
                if is_cpi_item:
                    matched_row = row
                    is_perfect_match = True
                    status_text = "정상"
                    break
                # 💡 [수정 여기까지]
            
            for _, row in df_item.iterrows():
                perfect = True
                row_all_text = "".join([str(v) for v in row.values]).replace(" ", "")
                row_vigo = str(row["비고"]) if "비고" in row else ""
                
                if "구동방식" in row and pd.notna(row["구동방식"]) and str(row["구동방식"]).strip() not in ["", "nan"]:
                    if str(row["구동방식"]).strip() != spec_구동: perfect = False
                if "인승" in row and pd.notna(row["인승"]) and str(row["인승"]).strip() not in ["", "nan"]:
                    if spec_인승 not in parse_cell_to_digits_list(row["인승"]): perfect = False
                if "속도" in row and pd.notna(row["속도"]) and str(row["속도"]).strip() not in ["", "nan"]:
                    if spec_속도 not in parse_cell_to_digits_list(row["속도"]): perfect = False
                if "도어오프닝" in row and pd.notna(row["도어오프닝"]) and str(row["도어오프닝"]).strip() not in ["", "nan"]:
                    if spec_오프닝 not in parse_cell_to_digits_list(row["도어오프닝"]): perfect = False
                
                if target_panel_col and target_panel_col in row and pd.notna(row[target_panel_col]) and str(row[target_panel_col]).strip() not in ["", "nan"]:
                    if str(row[target_panel_col]).strip() != spec_판넬: perfect = False
                    
                if "가이드" in row and pd.notna(row["가이드"]) and str(row["가이드"]).strip() not in ["", "nan"]:
                    if str(row["가이드"]).strip() != spec_가이드: perfect = False
                
                if any(opb in str(item) for opb in ["OPB", "MAIN", "SUB", "장애인", "CPI"]):
                    if not match_floor_range_from_vigo(row_vigo, stop_input):
                        perfect = False
                        
                if "웨이트" in str(item).replace(" ", ""):
                    if type_input not in row_all_text:
                        perfect = False
                
                if perfect:
                    matched_row = row
                    is_perfect_match = True
                    status_text = "정상"
                    break
            
            if not is_perfect_match:
                for _, row in df_item.iterrows():
                    score = 0
                    current_row_insung = 0
                    current_row_panel = ""
                    row_all_text = "".join([str(v) for v in row.values]).replace(" ", "")
                    row_vigo = str(row["비고"]) if "비고" in row else ""
                    
                    if "웨이트" in str(item).replace(" ", ""):
                        if type_input not in row_all_text: score += 20000
                    
                    if any(opb in str(item) for opb in ["OPB", "MAIN", "SUB", "장애인", "CPI"]):
                        if not match_floor_range_from_vigo(row_vigo, stop_input): score += 15000

                    if target_panel_col and target_panel_col in row and pd.notna(row[target_panel_col]) and str(row[target_panel_col]).strip() not in ["", "nan"]:
                        current_row_panel = str(row[target_panel_col]).strip()
                        if current_row_panel != spec_판넬: score += 10000
                    
                    if "구동방식" in row and pd.notna(row["구동방식"]) and str(row["구동방식"]).strip() not in ["", "nan"]:
                        if str(row["구동방식"]).strip() != spec_구동: score += 5000
                        
                    if "인승" in row and pd.notna(row["인승"]) and str(row["인승"]).strip() not in ["", "nan"]:
                        dig_list = [int(d) for d in parse_cell_to_digits_list(row["인승"])]
                        if dig_list:
                            diffs = [abs(int(spec_인승) - d) for d in dig_list]
                            score += min(diffs) * 10
                            current_row_insung = dig_list[0]
                            
                    if "속도" in row and pd.notna(row["속도"]) and str(row["속도"]).strip() not in ["", "nan"]:
                        dig_list = [int(d) for d in parse_cell_to_digits_list(row["속도"])]
                        if dig_list:
                            diffs = [abs(int(spec_속도) - d) for d in dig_list]
                            score += min(diffs) * 2
                            
                    if score < min_score:
                        min_score = score
                        best_fallback_row = row
                        p_label = f"({current_row_panel})" if current_row_panel else ""
                        status_text = f"⚠️ 근사치({current_row_insung}인승{p_label}) 대체" if current_row_insung else "⚠️ 주변사양 대체"
                
                matched_row = best_fallback_row
                
            raw_bom_rows.append({
                "품목명": item,
                "matched_row": matched_row,
                "status_text": status_text,
                "is_perfect_match": is_perfect_match
            })

        is_roller_mode = False
        for entry in raw_bom_rows:
            if "가이드슈" in str(entry["품목명"]).replace(" ", "") and entry["matched_row"] is not None:
                row_all_text = "".join([str(v) for v in entry["matched_row"].values]).replace(" ", "")
                if "롤러" in row_all_text:
                    is_roller_mode = True
                    break

        bom_results = []
        is_fire_door_selected = (fire_input == "방화도어")
        
        for entry in raw_bom_rows:
            item = entry["품목명"]
            matched_row = entry["matched_row"]
            status_text = entry["status_text"]
            is_perfect_match = entry["is_perfect_match"]
            
            if matched_row is not None:
                calc_type = str(matched_row["계산방식"]).strip() if "계산방식" in matched_row else "기본(1식)"
                try:
                    unit_price = int(float(str(matched_row["단가"]).replace(",", "").split(".")[0].strip()))
                except Exception:
                    unit_price = 0
                
                qty = parse_qty_from_note(
                    item, calc_type, matched_row.get("비고", ""), 
                    stop_input, floor_input, type_input, spec_구동, spec_인승,
                    is_roller_mode=is_roller_mode, is_fire_door=is_fire_door_selected
                )
                
                total_price = int(qty * unit_price)
                
                prod_name = str(matched_row["제품명"]).strip() if "제품명" in matched_row and pd.notna(matched_row["제품명"]) else ""
                if prod_name in ["nan", "None"]: prod_name = ""
                
                excel_vigo = str(matched_row.get("비고", "")).strip() if pd.notna(matched_row.get("비고", "")) else ""
                if excel_vigo in ["nan", "None"]: excel_vigo = ""
                
                item_noclean = str(item).replace(" ", "")
                
                # 비고 설명란 갱신
                if "케이블_외" in item_noclean or "케이블외" in item_noclean:
                    excel_vigo = f"[사양 강제 고정] 기본(1식) 반영 수량 1 고정 ({excel_vigo})".strip()
                elif "MAIN_OPB" in item_noclean:
                    excel_vigo = f"[{stop_input} STOP 구간 단가 매칭] 수량 1대 고정 ({excel_vigo})".strip()
                elif "CPI" in item_noclean:
                    excel_vigo = f"[{type_input} 사양 반영] 수량 {int(qty)}대 자동 연동 ({excel_vigo})".strip()
                elif "SUB_OPB" in item_noclean:
                    if type_input == "일방":
                        excel_vigo = f"[일방 사양 반영] 수량 0 대 표기 ({excel_vigo})".strip()
                    else:
                        excel_vigo = f"[관통 사양 반영] [{stop_input} STOP 구간 단가 매칭] ({excel_vigo})".strip()
                elif "장애인용_OPB" in item_noclean:
                    excel_vigo = f"[{stop_input} STOP 구간 단가 매칭] 카 내 1대 고정 ({excel_vigo})".strip()
                elif "웨이트" in item_noclean:
                    excel_vigo = f"[{type_input} 사양 반영] 단가 자동 매칭 ({excel_vigo})".strip()
                elif "조속기" in item_noclean and "와이어" in item_noclean:
                    excel_vigo = f"[현장고정수식] ({floor_input}층 * 3M + 15M) * 2 = {int(qty)}M 산출 ({excel_vigo})".strip()
                elif "와이어" in item_noclean and "소켓" in item_noclean:
                    insung_val = int(spec_인승)
                    if spec_구동.upper() == "MR":
                        w_so_qty = 6 if insung_val == 8 else (8 if insung_val == 10 else 10)
                    else:
                        w_so_qty = 10 if insung_val <= 13 else (12 if insung_val <= 21 else 14)
                    excel_vigo = f"[고정 수식 기준] {spec_구동} 타입 {insung_val}인승 -> {int(w_so_qty)}개 고정 ({excel_vigo})".strip()
                elif "와이어" in item_noclean and "소켓" not in item_noclean:
                    insung_val = int(spec_인승)
                    w_cols = 4 if insung_val <= 12 else (5 if insung_val == 13 else 6)
                    excel_vigo = f"[현장고정수식] 1선당 {floor_input}층*3M+9M = {(floor_input*3)+9}M | {w_cols}열 반영 ({excel_vigo})".strip()
                elif "가이드레일" in item_noclean:
                    excel_vigo = f"[현장고정수식] 층고 3M, OH+PIT 6M, 5M 레일 기준 올림 산출 ({excel_vigo})".strip()
                elif "레일조인트" in item_noclean or "조인트" in item_noclean:
                    excel_vigo = f"[현장고정수식] 레일 본수 - 1 적용 ({excel_vigo})".strip()
                
                if ("오일통" in item or "오일받이" in item) and is_roller_mode:
                    excel_vigo = "[롤러 사양 수량 0 제외] " + excel_vigo
                
                if "방화도어" in item_noclean and not is_fire_door_selected:
                    excel_vigo = "[일반도어 선택으로 수량 0 제외] " + excel_vigo
                
                bom_results.append({
                    "구분": matched_row.get("구분", "기타"),
                    "상태": status_text,
                    "품목명": item,
                    "제품명": prod_name,
                    "계산방식": "기본(1식)" if "케이블" in item_noclean else ("고정 수식" if any(x in item_noclean for x in ["가이드레일", "조인트", "와이어", "OPB", "장애인용", "CPI"]) else calc_type),
                    "계산수량": qty,
                    "단가(원)": unit_price,
                    "합계금액(원)": total_price,
                    "비고": excel_vigo,
                    "정상매칭여부": is_perfect_match 
                })
            else:
                bom_results.append({
                    "구분": "기타",
                    "상태": "❌ 누락",
                    "품목명": item,
                    "제품명": "데이터 없음",
                    "계산방식": "기본(1식)",
                    "계산수량": 0.0,
                    "단가(원)": 0,
                    "합계금액(원)": 0,
                    "비고": "단가표에 자재 데이터가 전무합니다.",
                    "정상매칭여부": False
                })
        
        if bom_results:
            st.subheader("📋 선택 사양 기준 자재 내역서 (BOM)")
            df_display = pd.DataFrame(bom_results)
            
            def style_rows(row):
                if not row['정상매칭여부']:
                    return ['background-color: #FFF2CC; color: #7F6000;'] * len(row)
                if row['계산수량'] == 0:
                    return ['color: #A0A0A0; font-style: italic; background-color: #F9F9F9;'] * len(row)
                return [''] * len(row)
            
            cols_order = ["구분", "상태", "품목명", "제품명", "계산방식", "계산수량", "단가(원)", "합계금액(원)", "비고"]
            df_styled = df_display.style.apply(style_rows, axis=1).format({
                '계산수량': '{:,.1f}',
                '단가(원)': '{:,.0f}',
                '합계금액(원)': '{:,.0f}'
            })
            
            st.dataframe(df_styled, use_container_width=True, column_order=cols_order)
            
            st.markdown("---")
            st.subheader("📁 파트별 원가 요약")
            df_part_summary = df_display.groupby("구분")["합계금액(원)"].sum().reset_index()
            df_part_summary.columns = ["구분 (파트)", "파트별 합계 금액 (원)"]
            st.dataframe(df_part_summary.style.format({'파트별 합계 금액 (원)': '{:,.0f}'}), use_container_width=True)
            
            total_sum = df_display["합계금액(원)"].sum()
            st.markdown("---")
            st.metric(label="💰 최종 산출 총 원가 (미일치 대체 항목 포함)", value=f"{total_sum:,} 원")
        else:
            st.warning("단가표 데이터가 비어 있습니다.")