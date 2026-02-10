import os
import msvcrt
import re
import win32com.client as win32
import time

def process_target_folder(excel, folder_path):
    """실제 데이터를 수집하고 요약 파일을 만드는 핵심 로직"""
    path_lower = folder_path.lower()
    is_arrhenius = "arrhenius plot" in path_lower
    is_humidity = "humidity" in path_lower
    
    # 두 키워드 중 어느 것도 해당하지 않으면 스킵
    if not (is_arrhenius or is_humidity):
        return

    data_list = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx") and not file.startswith("~$") and "_Summary" not in file:
                full_path = os.path.join(root, file)
                folder_name = os.path.basename(root)
                
                extracted_num = 0
                if is_arrhenius:
                    match = re.search(r'(\d+)\s*oc', folder_name, re.IGNORECASE)
                    extracted_num = int(match.group(1)) if match else 0
                else:
                    match = re.search(r'(\d+)\s*%', folder_name)
                    extracted_num = int(match.group(1)) if match else 0
                
                data_list.append((extracted_num, full_path))

    if not data_list:
        return

    data_list.sort(key=lambda x: x[0])
    extracted_values = []
    
    print(f"   📊 Processing: {os.path.basename(folder_path)} ({len(data_list)} files)")

    for i, (num, path) in enumerate(data_list, start=1):
        try:
            wb = excel.Workbooks.Open(os.path.abspath(path))
            if wb.Sheets.Count >= 3:
                ws = wb.Sheets(3)
                h2_val = ws.Range("H2").Value
                extracted_values.append((num, h2_val))
            wb.Close(False)
        except Exception as e:
            print(f"\n   ❌ Error reading {os.path.basename(path)}: {e}")
        
    if extracted_values:
        output_wb = excel.Workbooks.Add()
        output_ws = output_wb.ActiveSheet
        
        # --- [데이터 작성 및 수식 로직 (기존과 동일)] ---
        if is_arrhenius:
            headers = ["temp (℃)", "temp K", "1000/T (1000/K)", "σ (S/cm)", "ln(σ) (S/cm)"]
            for col, text in enumerate(headers, start=1):
                output_ws.Cells(1, col).Value = text
            for idx, (temp, sigma) in enumerate(extracted_values, start=2):
                output_ws.Cells(idx, 1).Value = temp
                output_ws.Cells(idx, 2).Formula = f"=A{idx}+273.15"
                output_ws.Cells(idx, 3).Formula = f"=1000/B{idx}"
                output_ws.Cells(idx, 4).Value = sigma
                if sigma and sigma > 0:
                    output_ws.Cells(idx, 5).Formula = f"=LN(D{idx})"
            last_col_letter, x_title, y_title = "E", "1000/T (1000/K)", "ln (σ)"
            chart_data_range = output_ws.Range(f"C1:C{len(extracted_values)+1},E1:E{len(extracted_values)+1}")
        else:
            output_ws.Cells(1, 1).Value = "RH (%)"
            output_ws.Cells(1, 2).Value = "σ (S/cm)"
            for idx, (rh, sigma) in enumerate(extracted_values, start=2):
                output_ws.Cells(idx, 1).Value = rh
                output_ws.Cells(idx, 2).Value = sigma
            last_col_letter, x_title, y_title = "B", "RH (%)", "σ (S/cm)"
            chart_data_range = output_ws.Range(output_ws.Cells(1, 1), output_ws.Cells(len(extracted_values)+1, 2))

        # 스타일 및 차트 생성 (기존 로직 유지)
        output_ws.Rows(1).Font.Bold = True
        output_ws.Columns(f"A:{last_col_letter}").HorizontalAlignment = -4108
        output_ws.Columns(f"A:{last_col_letter}").AutoFit()
        output_ws.Columns("H").HorizontalAlignment = -4108
        output_ws.Columns("H").AutoFit()

        chart_obj = output_ws.ChartObjects().Add(Left=60, Top=150, Width=450, Height=300)
        chart = chart_obj.Chart
        chart.ChartType = 74 
        chart.SetSourceData(Source=chart_data_range)
        chart.HasLegend = False
        
        series = chart.SeriesCollection(1)
        series.MarkerStyle = 8
        series.Format.Line.Visible = True
        series.Format.Line.Weight = 1.5
        series.Format.Line.ForeColor.ObjectThemeColor = 5

        if is_arrhenius:
            trendline = series.Trendlines().Add(Type=-4132)
            trendline.DisplayEquation = True
            time.sleep(0.3)
            try:
                eq = trendline.DataLabel.Text
                match = re.search(r'([-+]?\d*\.?\d+)(?:x)', eq.replace(" ", ""))
                slope_val = float(match.group(1)) if match else 0
            except: slope_val = 0

            # H열 계산 섹션
            output_ws.Range("H1").Value = "kB"
            output_ws.Range("H1").Font.Bold = True
            output_ws.Range("H2").Value = 0.08617333262
            output_ws.Range("H4").Value = "Slope"
            output_ws.Range("H4").Font.Bold = True
            output_ws.Range("H5").Value = slope_val
            output_ws.Range("H7").Value = "Activation energy (eV) = slope*kB"
            output_ws.Range("H7").Font.Bold = True
            output_ws.Range("H8").Formula = "=H5*H2"
            output_ws.Columns("H").AutoFit()

        for axis_type, title_text in zip([1, 2], [x_title, y_title]):
            axis = chart.Axes(axis_type)
            axis.HasTitle, axis.AxisTitle.Text = True, title_text
            axis.AxisTitle.Font.Size, axis.AxisTitle.Font.Bold = 15, True
            axis.HasMajorGridlines = False

        save_name = "Summary_Arrhenius plot.xlsx" if is_arrhenius else "Summary_Humidity dependent.xlsx"
        output_path = os.path.join(folder_path, save_name)
        output_wb.SaveAs(os.path.abspath(output_path))
        output_wb.Close()
        print(f"      ✅ Saved: {save_name}")

def main():
    while True:
        root_input = input("\n📂 Enter the root folder path (yyyy.mm.dd): ").strip().strip('"')
        
        if not os.path.isdir(root_input):
            print("❌ Invalid path. Please try again.")
            continue

        # 엑셀 실행 (한 번만)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            # 상위 폴더 내의 1단계 하위 폴더들 확인
            subfolders = [os.path.join(root_input, f) for f in os.listdir(root_input) 
                          if os.path.isdir(os.path.join(root_input, f))]
            
            found_any = False
            for folder in subfolders:
                folder_name_lower = os.path.basename(folder).lower()
                if "arrhenius plot" in folder_name_lower or "humidity dependent" in folder_name_lower:
                    process_target_folder(excel, folder)
                    found_any = True
            
            if not found_any:
                print("⚠️ No 'Arrhenius plot' or 'Humidity dependent' folders found.")
            else:
                print("\n✨ All tasks completed for the root folder!")

        finally:
            excel.Quit()

        print("\nPress Ctrl+F to process another root folder")
        print("Press any key to exit")
        if msvcrt.getch() != b'\x06':
            break

if __name__ == "__main__":
    main()