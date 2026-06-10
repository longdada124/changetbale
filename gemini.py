import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="課表彙整與代調課管理系統", layout="wide")

# --- 功能 1：密碼保護機制 ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🔐 課表彙整系統 - 系統登入</h2>", unsafe_allow_html=True)
    col_l, col_m, col_r = st.columns([1, 1.5, 1])
    with col_m:
        password = st.text_input("請輸入系統驗證密碼：", type="password")
        if st.button("確認登入", use_container_width=True):
            if password == "1030018":
                st.session_state.authenticated = True
                st.success("🎉 登入成功！")
                st.rerun()
            else:
                st.error("❌ 密碼錯誤，請重新輸入。")
    st.stop()

# --- 核心工具函數：動態產生20週日期與星期對照 ---
def generate_weeks(start_date_str, total_weeks=20):
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    # 強制調整到該週的週一
    start_date = start_date - timedelta(days=start_date.weekday())
    weeks_info = {}
    day_names = ["一", "二", "三", "四", "五"]
    
    for w in range(1, total_weeks + 1):
        w_start = start_date + timedelta(weeks=w-1)
        w_end = w_start + timedelta(days=4)
        label = f"第{w:02d}週 ({w_start.strftime('%m/%d')} ~ {w_end.strftime('%m/%d')})"
        
        days_map = {}
        date_to_day_num = {}
        for idx, name in enumerate(day_names):
            current_day = w_start + timedelta(days=idx)
            date_str = current_day.strftime("%Y-%m-%d")
            short_date_str = current_day.strftime("%m/%d") # 例如 07/20
            days_map[name] = {"date": date_str, "short": short_date_str}
            date_to_day_num[date_str] = {"day_idx": idx + 1, "day_name": name, "short": short_date_str}
            
        weeks_info[w] = {
            "label": label,
            "start_date": w_start,
            "days": days_map,
            "date_to_day_num": date_to_day_num
        }
    return weeks_info

# --- 核心 Word 替換函數 ---
def master_replace(doc_obj, old_text, new_text):
    if isinstance(new_text, (float, int)):
        new_val = str(int(new_text))
    else:
        new_val = str(new_text) if (new_text and str(new_text).strip() != "") else ""
    targets = list(doc_obj.paragraphs)
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                targets.extend(cell.paragraphs)
    for p in targets:
        if old_text in p.text:
            full_text = "".join([run.text for run in p.runs])
            updated_text = full_text.replace(old_text, new_val)
            for i, run in enumerate(p.runs):
                run.text = updated_text if i == 0 else ""

def load_default_template(file_name):
    try:
        with open(file_name, "rb") as f:
            return f.read()
    except FileNotFoundError:
        return None

# --- 側邊欄：基礎資料管理 ---
with st.sidebar:
    st.header("⚙️ 資料管理")
    if st.button("🔒 安全登出系統", type="secondary", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
        
    st.divider()
    if st.button("🧹 清空重置系統", use_container_width=True):
        for key in list(st.session_state.keys()): 
            if key != "authenticated": del st.session_state[key]
        st.rerun()

    st.divider()
    st.subheader("📅 學期開學日設定")
    semester_start = st.date_input("設定第一週週一的日期：", datetime(2026, 2, 16))

    st.divider()
    st.subheader("📥 範本下載")
    data_templates = {
        "1. 配課表範本": "配課表.xlsx",
        "2. 課表範本": "課表.xlsx",
        "3. 教師排序表範本": "教師排序表.xlsx"
    }
    for label, file_name in data_templates.items():
        try:
            with open(file_name, "rb") as f:
                st.download_button(label=f"{label}", data=f, file_name=file_name, key=f"dl_{file_name}")
        except FileNotFoundError:
            st.caption(f"⚠️ 找不到 {file_name}")
            
    st.divider()
    st.subheader("📤 上傳資料檔")
    f_assign = st.file_uploader("1. 上傳【配課表】", type=["xlsx", "csv"])
    f_time = st.file_uploader("2. 上傳【課表】", type=["xlsx", "csv"])
    f_sort = st.file_uploader("3. 上傳【教師排序暨時數表】", type=["xlsx", "csv"])
    
    if f_assign and f_time and st.button("🚀 執行整合並建構多週次課表", use_container_width=True):
        class_temp = load_default_template("班級樣板.docx")
        teacher_temp = load_default_template("教師樣板.docx")
        
        # 讀取增設的代調課樣板（若無則提示，但不中斷系統）
        sub_temp = load_default_template("代課樣板.docx")
        exc_temp = load_default_template("調課樣板.docx")
        
        if not class_temp or not teacher_temp:
            st.error("❌ 系統錯誤：後台找不到「班級樣板.docx」或「教師樣板.docx」。")
        else:
            with st.spinner("同步解析基礎資料並派發至20個學期週次中..."):
                df_assign = pd.read_csv(f_assign) if f_assign.name.endswith('.csv') else pd.read_excel(f_assign)
                df_time = pd.read_csv(f_time) if f_time.name.endswith('.csv') else pd.read_excel(f_time)
                
                # 初始化20週日期軸
                weeks_db = generate_weeks(semester_start.strftime("%Y-%m-%d"), total_weeks=20)
                st.session_state.weeks_db = weeks_db
                st.session_state.class_template = class_temp
                st.session_state.teacher_template = teacher_temp
                st.session_state.sub_template = sub_temp
                st.session_state.exc_template = exc_temp

                # 1. 解析配課
                assign_lookup, all_teachers_db, tutors = [], set(), {}
                for _, row in df_assign.iterrows():
                    c, s, t_raw = str(row['班級']).strip(), str(row['科目']).strip(), str(row['教師']).strip()
                    t_list = [name.strip() for name in t_raw.split('/')]
                    for t in t_list:
                        if t and t != "nan":
                            assign_lookup.append({'c': c, 's': s, 't': t})
                            all_teachers_db.add(t)
                    if s == "班級": tutors[c] = t_raw

                # 2. 教師排序與基礎授課時數
                ordered_teachers, base_hours, all_teachers_list = [], {}, list(all_teachers_db)
                if f_sort:
                    df_s = pd.read_csv(f_sort) if f_sort.name.endswith('.csv') else pd.read_excel(f_sort)
                    for _, s_row in df_s.iterrows():
                        t_name = str(s_row.iloc[0]).strip()
                        if t_name in all_teachers_list:
                            ordered_teachers.append(t_name)
                            try: base_hours[t_name] = int(s_row.iloc[1])
                            except: base_hours[t_name] = 0
                    for t in all_teachers_list:
                        if t not in ordered_teachers: ordered_teachers.append(t); base_hours[t] = 0
                else:
                    ordered_teachers = sorted(all_teachers_list)
                    base_hours = {t: 0 for t in ordered_teachers}

                # 3. 解析基礎原始課表
                base_class_data, base_teacher_data = {}, {}
                day_map = {"一":1,"二":2,"三":3,"四":4,"五":5,"週一":1,"週二":2,"週三":3,"週四":4,"週五":5}
                for _, row in df_time.iterrows():
                    c_raw, s_raw = str(row['班級']).strip(), str(row['科目']).strip()
                    d, p_match = day_map.get(str(row['星期']).strip(), 0), re.search(r'\d+', str(row['節次']))
                    if not (p_match and d > 0): continue
                    p = int(p_match.group())

                    if not s_raw or s_raw == "nan" or s_raw == "":
                        display_t = ""
                        s_raw = ""
                        curr_t_list = []
                    else:
                        curr_t_list = [item['t'] for item in assign_lookup if item['c'] == c_raw and item['s'] == s_raw]
                        display_t = "/".join(curr_t_list) if curr_t_list else "未知教師"
                    
                    if c_raw not in base_class_data: base_class_data[c_raw] = {}
                    base_class_data[c_raw][(d, p)] = {"subj": s_raw, "teacher": display_t, "label": ""}
                    
                    for t in curr_t_list:
                        if t not in base_teacher_data: base_teacher_data[t] = {}
                        base_teacher_data[t][(d, p)] = {"subj": s_raw, "class": c_raw, "label": ""}

                # 4. 生成獨立的 20 週資料集 (方案二核心：解除週次綁定)
                schedule_by_week = {}
                for w in range(1, 21):
                    schedule_by_week[w] = {"class_data": {}, "teacher_data": {}}
                    for c_key, c_val in base_class_data.items():
                        schedule_by_week[w]["class_data"][c_key] = {k: v.copy() for k, v in c_val.items()}
                    for t_key, t_val in base_teacher_data.items():
                        schedule_by_week[w]["teacher_data"][t_key] = {k: v.copy() for k, v in t_val.items()}

                st.session_state.update({
                    "schedule_by_week": schedule_by_week, "tutors_map": tutors, "base_hours": base_hours,
                    "ordered_teachers": ordered_teachers, "sel_class": sorted(list(base_class_data.keys()))[0] if base_class_data else "",
                    "sel_teacher": ordered_teachers[0] if ordered_teachers else ""
                })
                st.rerun()

# --- 主畫面大看板：學期與週次切換控制列 ---
if 'schedule_by_week' in st.session_state:
    st.markdown("### 🏫 學校課表動態調度與管理系統")
    
    # 完美複刻圖片頂部篩選列
    sel_col1, sel_col2, sel_col3 = st.columns([1, 1.5, 1])
    with sel_col1:
        st.selectbox("學年學期", ["114學年度 下學期"], index=0, disabled=True)
    with sel_col2:
        week_options = {w: info["label"] for w, info in st.session_state.weeks_db.items()}
        # 預設選取第 15 週
        current_w = st.selectbox("週次切換", options=list(week_options.keys()), format_func=lambda x: week_options[x], index=14)
    with sel_col3:
        st.caption("<p style='text-align:right; color:gray; padding-top:25px;'>系統運作正常 🟢</p>", unsafe_allow_html=True)

    # 動態計算「當週」每位老師授課實際總時數 (因應代調課會加減時數)
    current_total_counts = {t: 0 for t in st.session_state.ordered_teachers}
    for t in st.session_state.ordered_teachers:
        t_lessons = st.session_state.schedule_by_week[current_w]["teacher_data"].get(t, {})
        for (d, p), info in t_lessons.items():
            if info["subj"] != "":
                current_total_counts[t] += 1

    tab1, tab2, tab3 = st.tabs(["🏫 班級課表預覽", "👩‍🏫 教師課表預覽", "🔄 代、調課作業面板"])

    # --- Tab 1: 班級課表 ---
    with tab1:
        classes = sorted(list(st.session_state.schedule_by_week[current_w]["class_data"].keys()))
        curr_c = st.session_state.get('sel_class', classes[0])
        if curr_c not in classes: curr_c = classes[0]
        
        col1, col2, col3 = st.columns([1, 2, 1])
        if col1.button("⬅️ 上一班", key="btn_prev_c"):
            st.session_state.sel_class = classes[(classes.index(curr_c) - 1) % len(classes)]; st.rerun()
        if col3.button("下一班 ➡️", key="btn_next_c"):
            st.session_state.sel_class = classes[(classes.index(curr_c) + 1) % len(classes)]; st.rerun()
        with col2: 
            st.session_state.sel_class = st.selectbox("選取班級", classes, index=classes.index(curr_c), key="select_c")
        
        target_c = st.session_state.sel_class
        st.info(f"📍 當前班級：{target_c} | 導師：{st.session_state.tutors_map.get(target_c, '未設定')} | 顯示週次：{week_options[current_w]}")
        
        # 繪製功課表表格 (加入代調課狀態高亮標記)
        c_preview = []
        w_days = st.session_state.weeks_db[current_w]["days"]
        for p in range(1, 9):
            row = {"節次": f"第 {p} 節"}
            for d_idx, d_name in enumerate(["一", "二", "三", "四", "五"]):
                d = d_idx + 1
                info = st.session_state.schedule_by_week[current_w]["class_data"][target_c].get((d, p))
                if info and info["subj"] != "":
                    label_str = f"\n{info['label']}" if info['label'] else ""
                    row[f"週{d_name}\n({w_days[d_name]['short']})"] = f"{info['subj']}\n({info['teacher']}){label_str}"
                else:
                    row[f"週{d_name}\n({w_days[d_name]['short']})"] = ""
            c_preview.append(row)
        st.table(pd.DataFrame(c_preview))

        # 下載當週班級課表
        if st.button(f"📥 下載 {target_c} 第 {current_w} 週專屬課表"):
            doc = Document(BytesIO(st.session_state.class_template))
            master_replace(doc, "{{CLASS}}", target_c)
            master_replace(doc, "{{TUTOR}}", st.session_state.tutors_map.get(target_c, "未設定")) 
            for d, p in [(d,p) for d in range(1,6) for p in range(1,9)]:
                v = st.session_state.schedule_by_week[current_w]["class_data"][target_c].get((d,p), {"subj":"","teacher":"","label":""})
                lbl = f" {v['label']}" if v['label'] else ""
                master_replace(doc, f"{{{{SD{d}P{p}}}}}", v['subj'])
                master_replace(doc, f"{{{{TD{d}P{p}}}}}", f"{v['teacher']}{lbl}")
            buf = BytesIO(); doc.save(buf)
            st.download_button(f"💾 儲存 {target_c} 課表", buf.getvalue(), f"{target_c}_第{current_w}週課表.docx")

    # --- Tab 2: 教師課表 ---
    with tab2:
        teachers = st.session_state.ordered_teachers
        curr_t = st.session_state.get('sel_teacher', teachers[0])
        if curr_t not in teachers: curr_t = teachers[0]
        
        colt1, colt2, colt3 = st.columns([1, 2, 1])
        if colt1.button("⬅️ 前一位", key="btn_prev_t"):
            st.session_state.sel_teacher = teachers[(teachers.index(curr_t) - 1) % len(teachers)]; st.rerun()
        if colt3.button("下一位 ➡️", key="btn_next_t"):
            st.session_state.sel_teacher = teachers[(teachers.index(curr_t) + 1) % len(teachers)]; st.rerun()
        with colt2: 
            st.session_state.sel_teacher = st.selectbox("選取教師", teachers, index=teachers.index(curr_t), key="select_t")

        target_t = st.session_state.sel_teacher
        base = int(st.session_state.base_hours.get(target_t, 0))
        total = current_total_counts.get(target_t, 0)
        
        m1, m2, m3 = st.columns(3)
        m1.metric("基本應授時數", f"{base} 節")
        m2.metric(f"第 {current_w} 週實際總時數", f"{total} 節")
        m3.metric("本週兼代課/超時數", f"{total-base} 節")
        
        t_prev = []
        for p in range(1, 9):
            row = {"節次": f"第 {p} 節"}
            for d_idx, d_name in enumerate(["一", "二", "三", "四", "五"]):
                d = d_idx + 1
                info = st.session_state.schedule_by_week[current_w]["teacher_data"].get(target_t, {}).get((d,p))
                if info and info["subj"] != "":
                    lbl_str = f" {info['label']}" if info['label'] else ""
                    row[f"週{d_name}\n({w_days[d_name]['short']})"] = f"{info['class']} {info['subj']}{lbl_str}"
                else:
                    row[f"週{d_name}\n({w_days[d_name]['short']})"] = ""
            t_prev.append(row)
        st.table(pd.DataFrame(t_prev))

    # --- Tab 3: 代、調課作業核心邏輯面板 ---
    with tab3:
        st.subheader("🔄 臨時請假代課與跨課調課處理")
        mode = st.radio("選擇作業類型：", ["1. 辦理臨時請假代課", "2. 辦理雙向課堂對調 (情境 B)"])
        st.divider()

        if mode == "1. 辦理臨時請假代課":
            st.write("👉 **功能說明**：選取請假老師受更動的課堂，指派 B 老師代課。系統會自動扣除原任課老師當週時數、增加代課老師時數，並在主課表畫面上標註 `[代]`。")
            
            with st.form("sub_form"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    sub_date = st.date_input("請假課程日期", datetime(2026, 5, 18))
                with c2:
                    sub_period = st.selectbox("請假節次", [f"第 {i} 節" for i in range(1, 9)], index=0)
                with c3:
                    sub_class = st.selectbox("發生班級", classes)
                
                b_teacher = st.selectbox("指派代課教師 (B老師)：", teachers)
                submit_sub = st.form_submit_button("🔥 確認執行代課變更並匯出通知單")
                
                if submit_sub:
                    date_str = sub_date.strftime("%Y-%m-%d")
                    # 搜尋該日期落在第幾週
                    target_w = None
                    for w_idx, w_val in st.session_state.weeks_db.items():
                        if date_str in w_val["date_to_day_num"]:
                            target_w = w_idx
                            break
                    
                    if not target_w:
                        st.error("❌ 錯誤：選擇的日期不在此學期 20 週時間軸內！")
                    else:
                        date_details = st.session_state.weeks_db[target_w]["date_to_day_num"][date_str]
                        d_num = date_details["day_idx"]
                        p_num = int(re.search(r'\d+', sub_period).group())
                        
                        # 撈取該格子原始課表資訊
                        orig_cell = st.session_state.schedule_by_week[target_w]["class_data"][sub_class].get((d_num, p_num))
                        if not orig_cell or orig_cell["subj"] == "":
                            st.error(f"❌ 錯誤：{date_str} {sub_period} {sub_class} 原本就沒有排課，無法辦理請假代課！")
                        else:
                            orig_subj = orig_cell["subj"]
                            a_teacher = orig_cell["teacher"] # 原請假老師 (A老師)
                            
                            # 執行方案二：動態寫回與連動修改
                            # 1. 修改班級課表資訊
                            st.session_state.schedule_by_week[target_w]["class_data"][sub_class][(d_num, p_num)] = {
                                "subj": orig_subj, "teacher": b_teacher, "label": "[代]"
                            }
                            # 2. 扣除原任課 A 老師當週該堂課的登記 (使時數統計同步扣除)
                            if a_teacher in st.session_state.schedule_by_week[target_w]["teacher_data"]:
                                st.session_state.schedule_by_week[target_w]["teacher_data"][a_teacher][(d_num, p_num)] = {
                                    "subj": "", "class": "", "label": ""
                                }
                            # 3. 增加新代課 B 老師當週該堂課的登記 (使時數統計同步增加)
                            if b_teacher not in st.session_state.schedule_by_week[target_w]["teacher_data"]:
                                st.session_state.schedule_by_week[target_w]["teacher_data"][b_teacher] = {}
                            st.session_state.schedule_by_week[target_w]["teacher_data"][b_teacher][(d_num, p_num)] = {
                                "subj": orig_subj, "class": sub_class, "label": "[代]"
                            }
                            
                            st.success(f"🎉 變更成功！第 {target_w} 週的主課表已動態更新。")
                            
                            # 套用 Word 樣板匯出
                            if st.session_state.sub_template:
                                doc = Document(BytesIO(st.session_state.sub_template))
                                # 在這裡你可以設計樣板專屬標籤做取代
                                master_replace(doc, "{{DATE}}", date_str)
                                master_replace(doc, "{{CLASS}}", sub_class)
                                master_replace(doc, "{{PERIOD}}", sub_period)
                                master_replace(doc, "{{SUBJECT}}", orig_subj)
                                master_replace(doc, "{{LEAVE_TEACHER}}", a_teacher)
                                master_replace(doc, "{{SUB_TEACHER}}", b_teacher)
                                
                                buf = BytesIO(); doc.save(buf)
                                st.download_button("📥 下載代課通知單.docx", buf.getvalue(), f"{date_str.replace('-','')}_{sub_class}_代課單.docx")
                            else:
                                st.warning("💡 提示：主課表已修正。若需要下載通知單，請先在後台資料夾中放置「代課樣板.docx」。")

        elif mode == "2. 辦理雙向課堂對調 (情境 B)":
            st.write("👉 **功能說明**：設定兩門課務互相對調。系統會自動把這兩堂課的時間、班級、科目及老師全要素互換，並在主課表畫面上高亮標註 `[調 日期 星期-節次]`。")
            
            with st.form("exc_form"):
                st.markdown("##### 📍 課堂 A (原課務)")
                xa1, xa2, xa3 = st.columns(3)
                with xa1: date_a = st.date_input("課堂 A 日期", datetime(2026, 5, 21))
                with xa2: period_a = st.selectbox("課堂 A 節次", [f"第 {i} 節" for i in range(1, 9)], index=2) # 預設第3節
                with xa3: class_a = st.selectbox("課堂 A 班級", classes, key="ex_ca")
                
                st.markdown("##### 📍 課堂 B (欲對調之新課務)")
                xb1, xb2, xb3 = st.columns(3)
                with xb1: date_b = st.date_input("課堂 B 日期", datetime(2026, 5, 22))
                with xb2: period_b = st.selectbox("課堂 B 節次", [f"第 {i} 節" for i in range(1, 9)], index=4) # 預設第5節
                with xb3: class_b = st.selectbox("課堂 B 班級", classes, key="ex_cb")
                
                submit_exc = st.form_submit_button("🔥 確認執行雙向調課對調並匯出通知單")
                
                if submit_exc:
                    str_a, str_b = date_a.strftime("%Y-%m-%d"), date_b.strftime("%Y-%m-%d")
                    w_a, w_b = None, None
                    for w_idx, w_val in st.session_state.weeks_db.items():
                        if str_a in w_val["date_to_day_num"]: w_a = w_idx
                        if str_b in w_val["date_to_day_num"]: w_b = w_idx
                        
                    if not w_a or not w_b:
                        st.error("❌ 錯誤：選擇的日期超出學期 20 週的範圍！")
                    else:
                        dt_a = st.session_state.weeks_db[w_a]["date_to_day_num"][str_a]
                        dt_b = st.session_state.weeks_db[w_b]["date_to_day_num"][str_b]
                        p_a = int(re.search(r'\d+', period_a).group())
                        p_b = int(re.search(r'\d+', period_b).group())
                        
                        # 撈取兩格原始資料
                        cell_a = st.session_state.schedule_by_week[w_a]["class_data"][class_a].get((dt_a["day_idx"], p_a), {"subj":"","teacher":"","label":""}).copy()
                        cell_b = st.session_state.schedule_by_week[w_b]["class_data"][class_b].get((dt_b["day_idx"], p_b), {"subj":"","teacher":"","label":""}).copy()
                        
                        if cell_a["subj"] == "" and cell_b["subj"] == "":
                            st.error("❌ 錯誤：對調的兩堂課皆為空堂，無需調課！")
                        else:
                            # 產生符合你期望的標籤格式：[調 月/日 星期-節次]
                            label_for_a = f"[調 {dt_b['short']} {dt_b['day_name']}-{p_b}]"
                            label_for_b = f"[調 {dt_a['short']} {dt_a['day_name']}-{p_a}]"
                            
                            # 1. 班級課表雙向改寫
                            st.session_state.schedule_by_week[w_a]["class_data"][class_a][(dt_a["day_idx"], p_a)] = {
                                "subj": cell_b["subj"], "teacher": cell_b["teacher"], "label": label_for_a if cell_b["subj"] else ""
                            }
                            st.session_state.schedule_by_week[w_b]["class_data"][class_b][(dt_b["day_idx"], p_b)] = {
                                "subj": cell_a["subj"], "teacher": cell_a["teacher"], "label": label_for_b if cell_a["subj"] else ""
                            }
                            
                            # 2. 教師個人課表與時數結構同步雙向對調
                            t_a, t_b = cell_a["teacher"], cell_b["teacher"]
                            if t_a and t_a in st.session_state.schedule_by_week[w_a]["teacher_data"]:
                                st.session_state.schedule_by_week[w_a]["teacher_data"][t_a][(dt_a["day_idx"], p_a)] = {
                                    "subj": cell_b["subj"], "class": class_b if cell_b["subj"] else "", "label": label_for_a if cell_b["subj"] else ""
                                }
                            if t_b and t_b in st.session_state.schedule_by_week[w_b]["teacher_data"]:
                                st.session_state.schedule_by_week[w_b]["teacher_data"][t_b][(dt_b["day_idx"], p_b)] = {
                                    "subj": cell_a["subj"], "class": class_a if cell_a["subj"] else "", "label": label_for_b if cell_a["subj"] else ""
                                }
                                
                            st.success(f"🎉 調課對調完成！當週班級課表、教師課表、及兼代課時數已同步即時修正。")
                            
                            # 套用 Word 調課通知單匯出
                            if st.session_state.exc_template:
                                doc = Document(BytesIO(st.session_state.exc_template))
                                master_replace(doc, "{{ORG_DATE}}", str_a)
                                master_replace(doc, "{{ORG_WEEK}}", dt_a["day_name"])
                                master_replace(doc, "{{ORG_PERIOD}}", period_a)
                                master_replace(doc, "{{ORG_CLASS}}", class_a)
                                master_replace(doc, "{{ORG_TEACHER}}", t_a)
                                master_replace(doc, "{{ORG_SUBJECT}}", cell_a["subj"])
                                
                                master_replace(doc, "{{NEW_DATE}}", str_b)
                                master_replace(doc, "{{NEW_WEEK}}", dt_b["day_name"])
                                master_replace(doc, "{{NEW_PERIOD}}", period_b)
                                master_replace(doc, "{{NEW_CLASS}}", class_b)
                                master_replace(doc, "{{NEW_TEACHER}}", t_b)
                                master_replace(doc, "{{NEW_SUBJECT}}", cell_b["subj"])
                                
                                buf = BytesIO(); doc.save(buf)
                                st.download_button("📥 下載調課通知單.docx", buf.getvalue(), f"{str_a.replace('-','')}_調課單.docx")
                            else:
                                st.warning("💡 提示：主課表已修正。若需要下載通知單，請先在後台資料夾中放置「調課樣板.docx」。")
else:
    st.info("👋 請至側邊欄設定學期第一週日期、上傳三個基礎資料檔，並點擊「🚀 執行整合」以建構多週次動態調度系統。")