import os
import re
import time
import msvcrt
import win32com.client as win32
from openpyxl import load_workbook
from openpyxl.styles import Font

# --- [1. 개별 파일 수식 및 서식 설정] ---
def print_progress(current, total, bar_length=30):
    percent = current / total if total else 0
    filled = int(bar_length * percent)
    bar = "█" * filled + "-" * (bar_length - filled)
    print(f"\r   Progress: |{bar}| {percent*100:5.1f}% ({current}/{total})", end="")

def apply_formulas_to_all_files(root_path, l_value):
    xlsx_files = []
    for root, _, files in os.walk(root_path):
        for file in files:
            if file.endswith(".xlsx") and not file.startswith("~$") and "Summary" not in file:
                xlsx_files.append(os.path.join(root, file))

    total = len(xlsx_files)
    if total == 0: return

    print(f"📝 Step 1: Applying formulas to {total} files...")
    for i, path in enumerate(xlsx_files, start=1):
        try:
            wb = load_workbook(path)
            if wb.sheetnames:
                wb[wb.sheetnames[0]]["B5"].font = Font(bold=True)
            if len(wb.sheetnames) >= 3:
                ws3 = wb[wb.sheetnames[2]]
                last_row = ws3.max_row
                ws3.cell(row=1, column=5).value = "resistivity (ohm/cm)"
                ws3.cell(row=1, column=6).value = "conductivity (S/cm)"
                for r in range(2, last_row + 1):
                    ws3.cell(row=r, column=5).value = f"=(D{r}/C{r})*(2*PI()*{l_value})"
                    ws3.cell(row=r, column=6).value = f"=1/E{r}"
                ws3["H1"].value, ws3["H1"].font = "Average Conductivity", Font(bold=True)
                ws3["H2"].value, ws3["H2"].font = f"=AVERAGE(F3:F{last_row})", Font(bold=True)
            wb.save(path)
        except Exception as e:
            print(f"\n   ❌ Error in {os.path.basename(path)}: {e}")
        print_progress(i, total)
    print("\n   ✅ Step 1 Complete.")

# --- [2. 요약 파일 및 차트(Slope 추출) 생성] ---
def process_summary_folder(excel, folder_path):
    path_lower = folder_path.lower()
    is_arrhenius = "arrhenius plot" in path_lower
    is_humidity = "humidity" in path_lower
    if not (is_arrhenius or is_humidity): return

    data_list = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx") and not file.startswith("~$") and "Summary" not in file:
                f_name = os.path.basename(root)
                num = 0
                if is_arrhenius:
                    m = re.search(r'(\d+)\s*oc', f_name, re.IGNORECASE)
                    num = int(m.group(1)) if m else 0
                else:
                    m = re.search(r'(\d+)\s*%', f_name)
                    num = int(m.group(1)) if m else 0
                data_list.append((num, os.path.join(root, file)))

    if not data_list: return
    data_list.sort(key=lambda x: x[0])
    
    print(f"📊 Step 2: Creating summary for {os.path.basename(folder_path)}...")
    extracted = []
    for num, path in data_list:
        try:
            wb = excel.Workbooks.Open(os.path.abspath(path))
            val = wb.Sheets(3).Range("H2").Value
            extracted.append((num, val))
            wb.Close(False)
        except: pass
        
    if extracted:
        wb_out = excel.Workbooks.Add()
        ws_out = wb_out.ActiveSheet
        
        if is_arrhenius:
            for c, t in enumerate(["temp (℃)", "temp K", "1000/T", "σ (S/cm)", "ln(σ)"], 1): ws_out.Cells(1, c).Value = t
            for idx, (temp, sigma) in enumerate(extracted, 2):
                ws_out.Cells(idx, 1).Value = temp
                ws_out.Cells(idx, 2).Formula = f"=A{idx}+273.15"
                ws_out.Cells(idx, 3).Formula = f"=1000/B{idx}"
                ws_out.Cells(idx, 4).Value = sigma
                if sigma and sigma > 0: ws_out.Cells(idx, 5).Formula = f"=LN(D{idx})"
            last_c, x_t, y_t = "E", "1000/T (1000/K)", "ln (σ)"
            chart_range = ws_out.Range(f"C1:C{len(extracted)+1},E1:E{len(extracted)+1}")
        else:
            ws_out.Cells(1,1).Value, ws_out.Cells(1,2).Value = "RH (%)", "σ (S/cm)"
            for idx, (rh, sigma) in enumerate(extracted, 2):
                ws_out.Cells(idx,1).Value, ws_out.Cells(idx,2).Value = rh, sigma
            last_c, x_t, y_t = "B", "RH (%)", "σ (S/cm)"
            chart_range = ws_out.Range(ws_out.Cells(1,1), ws_out.Cells(len(extracted)+1, 2))

        # --- 가운데 정렬 및 스타일 설정 ---
        ws_out.Rows(1).Font.Bold = True
        ws_out.Columns(f"A:{last_c}").HorizontalAlignment = -4108 # 가운데 정렬
        ws_out.Columns(f"A:{last_c}").AutoFit()

        # 차트 생성
        chart_obj = ws_out.ChartObjects().Add(60, 150, 450, 300)
        chart = chart_obj.Chart
        chart.ChartType = 74 # xlXYScatterLines
        chart.SetSourceData(chart_range)
        chart.HasLegend = False
        
        slope = 0
        if is_arrhenius:
            series = chart.SeriesCollection(1)
            tl = series.Trendlines().Add(Type=-4132) # xlLinear
            tl.DisplayEquation = True
            
            # 수식 읽기 대기 로직 강화
            eq = ""
            for _ in range(15):
                time.sleep(0.3)
                try:
                    eq = tl.DataLabel.Text
                    if "x" in eq.lower() or "y" in eq.lower(): break
                except: continue
            
            if eq:
                # 공백 제거 및 x 앞의 숫자 추출 (한글 엑셀 'y = ' 대응)
                clean_eq = eq.replace(" ", "").replace(",", "")
                match = re.search(r'([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)x', clean_eq, re.IGNORECASE)
                if match:
                    slope = float(match.group(1))
                elif "=-x" in clean_eq.lower(): slope = -1.0
                elif "=x" in clean_eq.lower(): slope = 1.0

            # 결과 셀 정렬 및 출력
            ws_out.Range("H1").Value = "kB"
            ws_out.Range("H1").Font.Bold = True
            ws_out.Range("H2").Value = 0.08617333262
            ws_out.Range("H4").Value = "Slope"
            ws_out.Range("H4").Font.Bold = True
            ws_out.Range("H5").Value = slope
            ws_out.Range("H7").Value = "Activation energy (eV)"
            ws_out.Range("H7").Font.Bold = True
            ws_out.Range("H8").Formula = "=H5*H2"
            
            # H열도 가운데 정렬 및 자동 너비
            ws_out.Columns("H").HorizontalAlignment = -4108
            ws_out.Columns("H").AutoFit()

        for ax_t, t_txt in zip([1, 2], [x_t, y_t]):
            ax = chart.Axes(ax_t)
            ax.HasTitle, ax.AxisTitle.Text = True, t_txt
            ax.AxisTitle.Font.Size, ax.AxisTitle.Font.Bold = 14, True
            ax.HasMajorGridlines = False

        save_path = os.path.join(folder_path, f"Summary_{os.path.basename(folder_path)}.xlsx")
        wb_out.SaveAs(os.path.abspath(save_path))
        wb_out.Close()
        print(f"      ✅ Summary created. (Slope: {slope})")

# --- [3. 메인 실행 루프] ---
def main():
    while True:
        root_input = input("\n📂 Enter root folder: ").strip().strip('"')
        if not os.path.isdir(root_input):
            print("❌ Invalid path.")
            continue
        try:
            l_val = float(input("✍️ Enter l (cm): "))
        except:
            print("❌ Enter a number."); continue

        start = time.time()
        apply_formulas_to_all_files(root_input, l_val)

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible, excel.DisplayAlerts = False, False
        try:
            subs = [os.path.join(root_input, f) for f in os.listdir(root_input) if os.path.isdir(os.path.join(root_input, f))]
            for s in subs: process_summary_folder(excel, s)
        finally:
            excel.Quit()

        print(f"\n✨ Done! ({time.time()-start:.1f}s)")
        print("Press [Ctrl+F] to continue, any other key to exit.")
        if msvcrt.getch() != b'\x06': break

if __name__ == "__main__":
    main()