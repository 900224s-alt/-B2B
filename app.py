import streamlit as st
import pandas as pd
import io

# 設定網頁標題與圖示
st.set_page_config(page_title="B系統 - 軒郁/新普利 B2B 揀貨作業整合", layout="centered")

st.title("📦 B2B 揀貨資料整合系統 (B系統)")
st.write("請依序上傳對應的四個 Excel 檔案，系統將自動進行精準雙重比對並匯出整合總表。")

# ----------------- 資料清洗工具函式 ----------------- #
def clean_customer_name(name):
    name_str = str(name).upper()
    if 'MOMO' in name_str: return 'MOMO'
    if '屈臣氏' in str(name): return '屈臣氏'
    if '康是美' in str(name): return '康是美'
    if '寶雅' in str(name): return '寶雅'
    if 'PCHOME' in name_str: return 'PCHOME'
    return str(name)

def get_pallet_spec(customer_name):
    if customer_name == 'MOMO': return 'MOMO藍板'
    if customer_name == '寶雅': return '寶雅川字板'
    if customer_name == '屈臣氏': return '屈臣氏木板'
    if customer_name == '康是美': return '中華綠板'
    return '一般木板'

def get_owner(val):
    val_str = str(val).strip()
    if '2434723505' in val_str: return '軒郁'
    if '2797223005' in val_str: return '新普利'
    return val_str

# ----------------- 前端上傳介面 ----------------- #
task_file = st.file_uploader("1. 上傳【工作任務查詢】(Excel)", type=["xlsx", "xls"])
mgmt_file = st.file_uploader("2. 上傳【出庫管理】(Excel)", type=["xlsx", "xls"])
line_file = st.file_uploader("3. 上傳【出庫單行查詢】(Excel)", type=["xlsx", "xls"])
history_file = st.file_uploader("4. 上傳【交易歷史查詢】(Excel)", type=["xlsx", "xls"])

# 當四個檔案都上傳完成後，才顯示按鈕
if task_file and mgmt_file and line_file and history_file:
    if st.button("🚀 開始整合分析資料", type="primary"):
        with st.spinner("系統正在處理中，請稍候..."):
            try:
                # 讀取 Excel
                task_df = pd.read_excel(task_file)
                mgmt_df = pd.read_excel(mgmt_file)
                line_df = pd.read_excel(line_file)
                history_df = pd.read_excel(history_file)

                # 欄位名稱定義
                col_line_owner = '貨主'          
                col_line_out_order = '出庫單號'  
                col_line_oms_order = 'OMS訂單號' 
                col_line_barcode = '貨品'        
                col_line_name = '描述'           
                col_line_qty = '數量總計'        
                col_line_expiry = '效期(YYYY-MM-DD)' 
                col_line_batch = '批號'          

                col_mgmt_oms = 'OMS訂單號'       
                col_mgmt_cust = '發貨名稱'       
                col_mgmt_date = '希望配達日'     

                col_hist_ref = '參考單號'        
                col_hist_id = '內部鍵ID'         
                col_hist_barcode = '貨品'       

                col_task_id = '內部指令號碼'     
                col_task_loc = '從儲位'          
                col_task_pal = '工作單位'        

                # 步驟一：出庫單行 + 出庫管理
                df_line_mgmt = pd.merge(line_df, mgmt_df[[col_mgmt_oms, col_mgmt_cust, col_mgmt_date]], 
                                        left_on=col_line_oms_order, right_on=col_mgmt_oms, how='left')

                # 步驟二：交易歷史 + 工作任務
                df_hist_task = pd.merge(history_df[[col_hist_ref, col_hist_id, col_hist_barcode]], 
                                        task_df[[col_task_id, col_task_loc, col_task_pal]], 
                                        left_on=col_hist_id, right_on=col_task_id, how='left')

                # 步驟三：雙重鍵值精準合併
                df_final_raw = pd.merge(
                    df_line_mgmt, 
                    df_hist_task[[col_hist_ref, col_hist_barcode, col_task_loc, col_task_pal]], 
                    left_on=[col_line_out_order, col_line_barcode], 
                    right_on=[col_hist_ref, col_hist_barcode], 
                    how='left'
                )

                # 資料清洗與重組
                output_data = []
                for _, row in df_final_raw.iterrows():
                    raw_cust_name = row[col_mgmt_cust] if pd.notna(row[col_mgmt_cust]) else ""
                    clean_cust = clean_customer_name(raw_cust_name)

                    record = {
                        '貨主': get_owner(row.get(col_line_owner, "")),
                        '客戶名稱': clean_cust,
                        '訂單編號': row.get(col_line_out_order, ""),
                        '客戶訂單編號': row.get(col_line_oms_order, ""),
                        '希望配達日': row.get(col_mgmt_date, ""),
                        '儲位': row.get(col_task_loc, ""),
                        '棧板編號': row.get(col_task_pal, ""),
                        '棧板規格': get_pallet_spec(clean_cust),
                        '商品條碼': row.get(col_line_barcode, ""),
                        '品名': row.get(col_line_name, ""),
                        '數量': row.get(col_line_qty, ""),
                        '效期': row.get(col_line_expiry, ""),
                        '承運商': "",  
                        '批號': row.get(col_line_batch, "")
                    }
                    output_data.append(record)

                df_result = pd.DataFrame(output_data)

                # 排序：商品條碼 > 效期 > 批號 > 棧板編號
                df_result = df_result.sort_values(
                    by=['商品條碼', '效期', '批號', '棧板編號'], 
                    ascending=[True, True, True, True]
                )

                # 將結果寫入記憶體中的 Excel 檔 (讓使用者下載)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False)
                buffer.seek(0)

                st.success("🎉 資料整合成功！請點擊下方按鈕下載報表。")
                
                # 下載按鈕
                st.download_button(
                    label="💾 下載【軒郁_B2B揀貨整合總表.xlsx】",
                    data=buffer,
                    file_name="軒郁_B2B揀貨整合總表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 處理失敗。原因：{str(e)}")
                st.info("請檢查上傳的 Excel 檔案欄位標題是否與設定一致。")