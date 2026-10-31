# -*- coding: utf-8 -*-
"""Excelの「詳細データ」シートをTSVにするWindows DnD画面です。"""

import os
import subprocess
import sys
from typing import List

import win32api
import win32con
import win32gui


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
WINDOW_TITLE = "Group Shipment Summary (Drag & Drop)"


def show_message_box(pszMessage: str, pszTitle: str = WINDOW_TITLE) -> None:
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONINFORMATION)


def show_error_message_box(pszMessage: str, pszTitle: str = WINDOW_TITLE) -> None:
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONERROR)


def show_completion_dialog(
    pszOutputDirectory: str, bHasWarning: bool, pszDetails: str
) -> None:
    """短い完了画面と、スクロール可能な詳細画面を表示します。"""
    import tkinter as tk
    from tkinter import ttk

    iOwner = win32gui.GetForegroundWindow()
    objMonitor = win32api.MonitorFromWindow(iOwner, win32con.MONITOR_DEFAULTTONEAREST)
    iLeft, iTop, iRight, iBottom = win32api.GetMonitorInfo(objMonitor)["Work"]
    # ウィンドウ枠・タイトルバーも作業領域に収まる余白を確保します。
    iMaxWidth = max(1, iRight - iLeft - 64)
    iMaxHeight = max(1, iBottom - iTop - 96)
    objRoot = tk.Tk()
    objRoot.withdraw()
    objRoot.title(WINDOW_TITLE)
    objRoot.attributes("-topmost", True)
    bOwnerDisabled = False

    def place_window(objWindow: object, iWidth: int, iHeight: int) -> None:
        iWidth = min(iWidth, iMaxWidth)
        iHeight = min(iHeight, iMaxHeight)
        iX = iLeft + (iRight - iLeft - iWidth) // 2
        iY = iTop + (iBottom - iTop - iHeight) // 2
        objWindow.maxsize(iMaxWidth, iMaxHeight)
        objWindow.geometry(f"{iWidth}x{iHeight}")
        objWindow.update_idletasks()
        win32gui.SetWindowPos(
            win32gui.GetParent(objWindow.winfo_id()), 0, iX, iY, 0, 0,
            win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE,
        )

    def open_output_directory() -> None:
        try:
            if not os.path.isdir(pszOutputDirectory):
                raise OSError("output directory was not found")
            os.startfile(pszOutputDirectory)
        except OSError:
            show_error_message_box("出力フォルダーを開けませんでした。表示されたパスを確認してください。")

    def show_details() -> None:
        objDetails = tk.Toplevel(objRoot)
        objDetails.withdraw()
        objDetails.title("処理結果の詳細")
        objDetails.transient(objRoot)
        place_window(objDetails, 900, 560)
        objDetails.rowconfigure(0, weight=1)
        objDetails.columnconfigure(0, weight=1)
        objTextFrame = ttk.Frame(objDetails, padding=8)
        objTextFrame.grid(row=0, column=0, sticky="nsew")
        objTextFrame.rowconfigure(0, weight=1)
        objTextFrame.columnconfigure(0, weight=1)
        objText = tk.Text(objTextFrame, wrap="none", width=1, height=1)
        objVertical = ttk.Scrollbar(objTextFrame, orient="vertical", command=objText.yview)
        objHorizontal = ttk.Scrollbar(objTextFrame, orient="horizontal", command=objText.xview)
        objText.configure(yscrollcommand=objVertical.set, xscrollcommand=objHorizontal.set)
        objText.grid(row=0, column=0, sticky="nsew")
        objVertical.grid(row=0, column=1, sticky="ns")
        objHorizontal.grid(row=1, column=0, sticky="ew")
        objText.insert("1.0", pszDetails)
        objText.configure(state="disabled")

        def select_all(objEvent: object) -> str:
            objText.tag_add("sel", "1.0", "end-1c")
            return "break"

        objText.bind("<Control-a>", select_all)
        objText.bind("<Control-A>", select_all)
        objClose = ttk.Button(objDetails, text="閉じる", command=objDetails.destroy)
        objClose.grid(row=1, column=0, pady=(0, 8))
        objDetails.bind("<Escape>", lambda event: objDetails.destroy())
        objDetails.bind("<Return>", lambda event: objDetails.destroy())
        objDetails.protocol("WM_DELETE_WINDOW", objDetails.destroy)
        objDetails.deiconify()
        objDetails.grab_set()
        objText.focus_set()
        objRoot.wait_window(objDetails)
        objRoot.grab_set()
        objOk.focus_set()

    try:
        place_window(objRoot, 700, 300)
        objRoot.columnconfigure(0, weight=1)
        objRoot.rowconfigure(0, weight=1)
        objSummary = ttk.Frame(objRoot, padding=12)
        objSummary.grid(row=0, column=0, sticky="nsew")
        objSummary.columnconfigure(0, weight=1)
        ttk.Label(
            objSummary,
            text="処理が警告付きで完了しました。" if bHasWarning else "処理が完了しました。",
        ).grid(row=0, column=0, sticky="w", pady=(0, 12))
        ttk.Label(objSummary, text="出力フォルダー:").grid(row=1, column=0, sticky="w")
        objPathValue = tk.StringVar(objRoot, value=pszOutputDirectory)
        objPath = ttk.Entry(objSummary, state="readonly", textvariable=objPathValue)
        objPath.grid(row=2, column=0, sticky="ew", pady=4)
        objPathScroll = ttk.Scrollbar(objSummary, orient="horizontal", command=objPath.xview)
        objPath.configure(xscrollcommand=objPathScroll.set)
        objPathScroll.grid(row=3, column=0, sticky="ew")
        ttk.Label(objSummary, text="警告: " + ("あり" if bHasWarning else "なし")).grid(
            row=4, column=0, sticky="w", pady=8
        )
        objButtons = ttk.Frame(objRoot, padding=8)
        objButtons.grid(row=1, column=0, sticky="ew")
        arrButtons = [
            ttk.Button(objButtons, text="詳細", command=show_details),
            ttk.Button(objButtons, text="出力フォルダーを開く", command=open_output_directory),
            ttk.Button(objButtons, text="OK", command=objRoot.destroy),
        ]
        bStackButtons = iMaxWidth < 600
        for iIndex, objButton in enumerate(arrButtons):
            objButton.grid(row=iIndex if bStackButtons else 0, column=0 if bStackButtons else iIndex, padx=4, pady=3, sticky="ew")
            objButtons.columnconfigure(0 if bStackButtons else iIndex, weight=1)
        objOk = arrButtons[-1]
        objRoot.bind("<Return>", lambda event: objRoot.destroy())
        objRoot.bind("<Escape>", lambda event: objRoot.destroy())
        objRoot.protocol("WM_DELETE_WINDOW", objRoot.destroy)
        if iOwner and win32gui.IsWindow(iOwner) and win32gui.IsWindowEnabled(iOwner):
            win32gui.EnableWindow(iOwner, False)
            bOwnerDisabled = True
        objRoot.deiconify()
        objRoot.grab_set()
        objOk.focus_set()
        objRoot.mainloop()
    finally:
        if bOwnerDisabled and win32gui.IsWindow(iOwner):
            win32gui.EnableWindow(iOwner, True)
            win32gui.SetForegroundWindow(iOwner)
        try:
            objRoot.destroy()
        except tk.TclError:
            pass


def create_output_path(pszInputExcelPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputExcelPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszInputExcelPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszBaseName + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszBaseName}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def validate_dropped_file(pszInputExcelPath: str) -> str:
    if not os.path.isfile(pszInputExcelPath):
        return "Error: dropped item is not a file. Path = " + pszInputExcelPath
    if os.path.splitext(pszInputExcelPath)[1].lower() not in SUPPORTED_EXTENSIONS:
        return (
            "Error: unsupported Excel extension. Path = "
            + pszInputExcelPath
            + "\n\nSupported extensions: .xlsx, .xlsm"
        )
    return ""


def extract_error_file_path(pszStdErr: str) -> str:
    for pszLine in pszStdErr.splitlines():
        if pszLine.startswith("Error file: "):
            return pszLine[len("Error file: ") :].strip()
    return ""


def extract_warning_file_path(pszStdOut: str) -> str:
    for pszLine in pszStdOut.splitlines():
        if pszLine.startswith("警告ファイル: "):
            return pszLine[len("警告ファイル: ") :].strip()
    return ""


def run_excel_to_tsv(pszInputExcelPath: str) -> None:
    pszScriptPath: str = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "group_shipment_summary_Cmd.py"
    )
    if not os.path.isfile(pszScriptPath):
        show_error_message_box(
            "Error: group_shipment_summary_Cmd.py not found. Path = " + pszScriptPath
        )
        return

    pszValidationError: str = validate_dropped_file(pszInputExcelPath)
    if pszValidationError:
        show_error_message_box(pszValidationError)
        return

    pszOutputTsvPath: str = create_output_path(pszInputExcelPath)
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [sys.executable, pszScriptPath, pszOutputTsvPath, pszInputExcelPath],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as objException:
        show_error_message_box(
            "Error: failed to run converter. Detail = " + str(objException)
        )
        return

    if objCompletedProcess.returncode != 0 or not os.path.isfile(pszOutputTsvPath):
        pszStdErr: str = objCompletedProcess.stderr.strip() or "No stderr output."
        pszErrorFilePath: str = extract_error_file_path(pszStdErr)
        pszMessage: str = (
            "ExcelからTSVへの変換、年月別分割、グループ別集計・品種別集計・総計集計、または"
            "品種別step0003の43項目と農家廃棄羽数・農家廃棄重量・日令の3項目抽出、日令TSV複製による飼育日数平均ファイル作成、廃棄控除前・控除廃棄育成率作成、品種別step0004の6項目結合とstep0005年間連結・step0006年間行列入れ替え・step0007年間品種別集約、控除廃棄品種別step0004作成とstep0005年間連結・step0006年間行列入れ替え・step0007年間品種別集約、および全廃棄羽数・全廃棄重量・処理羽数・処理重量・控除後平均体重の5項目作成、最新5年のグループ別・総計43項目抽出、総計入雛羽数・平均体重結合、総計出荷羽数・出荷重量結合、総計羽数15項目の過去5年表作成とstep0005縦結合、総計廃棄羽数・重量各15項目結合と小計追加、懸鳥前羽数・重量各2項目結合と小計追加、step0006廃棄データ一覧表上段・下段作成、step0007同表上段・下段の合計列追加、step0008総計廃棄データ一覧表作成、step0009総計廃棄データ一覧表加工、step0010総計廃棄データ一覧表月別連結、グループ別4項目結合、または"
            "グループ別年間連結、年間TSVの行列入れ替え、または"
            "年別グループ連結・step0008年間TSV作成・年間合計追加・step0009年間TSV作成・年間合計列追加・step0013グループ別年間Excel作成・step0010グループ別年間TSVとstep0001日別稼動判定TSVから稼動日数行を追加したstep0011グループ別年間TSV作成・step0012前年実績追加・step0013前年比追加とグループ別年間Excel作成・step0014控除廃棄グループ別Excel作成・step0017控除廃棄グループ別Excel作成・step0015年間TSV作成・控除廃棄品種別step0008・step0009・step0010年間TSV作成・step0010年間Excel作成・品種別step0009・step0010年間TSV作成・step0010年間Excel作成に"
            "失敗しました。\n\nReturn code = "
            + str(objCompletedProcess.returncode)
            + "\n\nstderr:\n"
            + pszStdErr
        )
        if pszErrorFilePath:
            pszMessage += "\n\nエラー詳細:\n" + pszErrorFilePath
        show_error_message_box(pszMessage)
        return

    pszStdOut: str = objCompletedProcess.stdout.strip()
    pszWarningFilePath: str = extract_warning_file_path(pszStdOut)
    pszMessage = (
        (
            "処理が警告付きで完了しました。"
            if pszWarningFilePath
            else "処理が完了しました。"
        )
        + "\n\n入力Excel:\n"
        + pszInputExcelPath
        + "\n\n出力フォルダ:\n"
        + os.path.dirname(pszOutputTsvPath)
        + "\n\n元TSV:\n"
        + pszOutputTsvPath
    )
    arrDailyOperationPaths: List[str] = [
        pszLine[len("日別稼動判定ファイル: ") :]
        for pszLine in pszStdOut.splitlines()
        if pszLine.startswith("日別稼動判定ファイル: ")
    ]
    if arrDailyOperationPaths:
        pszMessage += "\n\n日別稼動判定TSV:\n" + "\n".join(arrDailyOperationPaths)
    arrStep0012VarietyPaths: List[str] = [
        line[len("step0012品種別年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0012品種別年間ファイル: ")
    ]
    if arrStep0012VarietyPaths:
        pszMessage += "\n\n前年実績付きstep0012 品種別TSV:\n" + "\n".join(arrStep0012VarietyPaths)
    arrStep0013VarietyPaths: List[str] = [
        line[len("step0013品種別年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0013品種別年間ファイル: ")
    ]
    if arrStep0013VarietyPaths:
        pszMessage += "\n\n一日平均出荷羽数の前年比付きstep0013 品種別TSV:\n" + "\n".join(arrStep0013VarietyPaths)
    arrStep0013VarietyExcelPaths: List[str] = [
        line[len("step0013品種別年間Excelファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0013品種別年間Excelファイル: ")
    ]
    if arrStep0013VarietyExcelPaths:
        pszMessage += "\n\nstep0013 品種別年間Excel:\n" + "\n".join(arrStep0013VarietyExcelPaths)
    arrNamedVarietySummaryExcelPaths: List[str] = [
        line[len("年付き品種別実績集計Excelファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("年付き品種別実績集計Excelファイル: ")
    ]
    if arrNamedVarietySummaryExcelPaths:
        pszMessage += "\n\n年付き品種別実績集計Excel:\n" + "\n".join(arrNamedVarietySummaryExcelPaths)
    arrStep0011VarietyDeductionPaths: List[str] = [
        line[len("step0011品種別控除廃棄年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0011品種別控除廃棄年間ファイル: ")
    ]
    if arrStep0011VarietyDeductionPaths:
        pszMessage += "\n\n稼動日数付きstep0011 品種別控除廃棄TSV:\n" + "\n".join(arrStep0011VarietyDeductionPaths)
    arrStep0012VarietyDeductionPaths: List[str] = [
        line[len("step0012品種別控除廃棄年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0012品種別控除廃棄年間ファイル: ")
    ]
    if arrStep0012VarietyDeductionPaths:
        pszMessage += "\n\n前年実績付きstep0012 品種別控除廃棄TSV:\n" + "\n".join(arrStep0012VarietyDeductionPaths)
    arrStep0013VarietyDeductionPaths: List[str] = [
        line[len("step0013品種別控除廃棄年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0013品種別控除廃棄年間ファイル: ")
    ]
    if arrStep0013VarietyDeductionPaths:
        pszMessage += "\n\n一日平均処理羽数の前年比付きstep0013 品種別控除廃棄TSV:\n" + "\n".join(arrStep0013VarietyDeductionPaths)
    arrStep0013VarietyDeductionExcelPaths: List[str] = [
        line[len("step0013品種別控除廃棄年間Excelファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0013品種別控除廃棄年間Excelファイル: ")
    ]
    if arrStep0013VarietyDeductionExcelPaths:
        pszMessage += "\n\nstep0013 品種別控除廃棄Excel:\n" + "\n".join(arrStep0013VarietyDeductionExcelPaths)
    arrNamedVarietyDeductionSummaryExcelPaths: List[str] = [
        line[len("年付き品種別控除廃棄実績集計Excelファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("年付き品種別控除廃棄実績集計Excelファイル: ")
    ]
    if arrNamedVarietyDeductionSummaryExcelPaths:
        pszMessage += "\n\n年付き品種別控除廃棄実績集計Excel:\n" + "\n".join(arrNamedVarietyDeductionSummaryExcelPaths)
    arrStep0012Paths: List[str] = [
        line[len("step0012グループ別年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0012グループ別年間ファイル: ")
    ]
    if arrStep0012Paths:
        pszMessage += "\n\n前年実績付きstep0012 TSV:\n" + "\n".join(arrStep0012Paths)
    arrStep0016Paths: List[str] = [
        line[len("step0016グループ別控除廃棄年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0016グループ別控除廃棄年間ファイル: ")
    ]
    if arrStep0016Paths:
        pszMessage += "\n\n前年実績付きstep0016 TSV:\n" + "\n".join(arrStep0016Paths)
    arrStep0017Paths: List[str] = [
        line[len("step0017グループ別控除廃棄年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0017グループ別控除廃棄年間ファイル: ")
    ]
    if arrStep0017Paths:
        pszMessage += "\n\n一日平均処理羽数の前年比付きstep0017 TSV:\n" + "\n".join(arrStep0017Paths)
    arrStep0013Paths: List[str] = [
        line[len("step0013グループ別年間ファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("step0013グループ別年間ファイル: ")
    ]
    if arrStep0013Paths:
        pszMessage += "\n\n一日平均出荷羽数の前年比付きstep0013 TSV:\n" + "\n".join(arrStep0013Paths)
    arrNamedGroupDeductionSummaryExcelPaths: List[str] = [
        line[len("年付きグループ別控除廃棄集計Excelファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("年付きグループ別控除廃棄集計Excelファイル: ")
    ]
    if arrNamedGroupDeductionSummaryExcelPaths:
        pszMessage += "\n\n年付きグループ別控除廃棄集計Excel:\n" + "\n".join(arrNamedGroupDeductionSummaryExcelPaths)
    arrNamedGroupSummaryExcelPaths: List[str] = [
        line[len("年付きグループ別集計Excelファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("年付きグループ別集計Excelファイル: ")
    ]
    if arrNamedGroupSummaryExcelPaths:
        pszMessage += "\n\n年付きグループ別集計Excel:\n" + "\n".join(arrNamedGroupSummaryExcelPaths)
    arrNamedTotalCountHistoryExcelPaths: List[str] = [
        line[len("出荷廃棄羽数過去5年集計Excelファイル: ") :]
        for line in pszStdOut.splitlines()
        if line.startswith("出荷廃棄羽数過去5年集計Excelファイル: ")
    ]
    if arrNamedTotalCountHistoryExcelPaths:
        pszMessage += "\n\n出荷廃棄羽数過去5年集計Excel:\n" + "\n".join(arrNamedTotalCountHistoryExcelPaths)
    if pszWarningFilePath:
        pszMessage += "\n\n警告ファイル:\n" + pszWarningFilePath
    show_completion_dialog(os.path.dirname(pszOutputTsvPath), bool(pszWarningFilePath), pszMessage)


def draw_instruction_text(iWindowHandle: int) -> None:
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    objClientRect = win32gui.GetClientRect(iWindowHandle)
    iMargin: int = 8
    objClientRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin,
    )
    pszInstructionText: str = (
        "Excelファイルをこのウィンドウに\n"
        "ドラッグ＆ドロップしてください。\n\n"
        "Excelの「詳細データ」シートをTSVへ変換します。\n"
        "処理対象の年月は入力データに含まれる年月から判定します。\n"
        "入力データを年月別のstep0001 TSVに分割します。\n"
        "年月別ファイルは入力データ内にある年月分を作成します。\n"
        "各月について、step0002のグループ別・品種別・総計TSVを作成します。\n"
        "step0001とstep0002の各ファイルは入力Excelと同じフォルダーに作成します。\n"
        "品種別集計は入力のP列（品種No.）と\n"
        "Q列（品種名）を使って行います。\n"
        "品種No.と品種名の組み合わせは次の4種類です。\n"
        "品種No.と品種名は両方とも一致する必要があります。\n"
        "どちらか一方だけが一致する場合もエラーになります。\n"
        "品種No. 5：阿波尾鶏\n"
        "品種No. 6：赤どり\n"
        "品種No. 2：地養鳥\n"
        "品種No. 1：ブロイラー\n"
        "step0002品種別TSVは上記4行の後に、ブロイラー（貞光分）、ブロイラー(外部購入)の順で計6行を出力します。\n"
        "該当する明細がない品種も行を出力します。\n"
        "明細がない品種の合計値は0を出力します。\n"
        "明細がない品種の平均体重と日齢は空欄です。\n"
        "既存の4品種行はグループを区別せずに合算します。\n"
        "ブロイラーの契約・直営は貞光分、フリーは外部購入として追加集計します。\n"
        "品種別TSVのK列とL列には「－」を出力します。\n"
        "貞光分・外部購入の振り分けには入力K列を使い、グループ別集計ではK列・L列の対応を検査します。\n"
        "品種別TSVのA列に年、B列に月を出力します。\n"
        "品種別TSVのD列に平均体重を出力します。\n"
        "品種別TSVのP列に品種No.、Q列に品種名を出力します。\n"
        "E列・F列は対象明細の値を合計します。\n"
        "O列には処理年月をYYYYMM形式で出力します。\n"
        "R列は入雛羽数、S列は出荷羽数を合計します。\n"
        "T列には日齢の単純平均を出力します。\n"
        "U列からBL列までは対象明細の値を合計します。\n"
        "上記に記載していない列には全角の「－」を出力します。\n"
        "平均体重は入力D列の平均ではありません。\n"
        "出荷重量合計を出荷羽数合計で割って再計算します。\n"
        "出荷羽数合計が0の場合、平均体重は空欄にします。\n"
        "出荷羽数がすべて空欄の場合も平均体重は空欄です。\n"
        "日齢は数値がある明細の単純平均を計算します。\n"
        "合計対象列は空欄を除いて合計します。\n"
        "合計対象列がすべて空欄の場合は0を出力します。\n"
        "合計値は不要な小数末尾のゼロを除いて出力します。\n"
        "平均体重は小数点以下5桁で四捨五入します。\n"
        "日齢は小数点以下2桁で四捨五入します。\n"
        "数値計算にはDecimalを使用します。\n"
        "出力TSVのヘッダーはstep0001のヘッダーを引き継ぎます。\n"
        "ヘッダーと明細の列数は64列である必要があります。\n"
        "入力に明細行がない場合はエラーになります。\n"
        "1つのstep0001ファイルに複数年月がある場合はエラーです。\n"
        "年・月または処理日付が不正な場合はエラーです。\n"
        "A列・B列の年月とO列の日付の年月が違う場合はエラーです。\n"
        "品種No.と品種名が指定の対応に一致しない場合はエラーです。\n"
        "数値列または日齢に空欄以外の非数値があればエラーです。\n"
        "合計対象列がすべて空欄の場合は警告を記録します。\n"
        "日齢がすべて空欄の場合は警告を記録します。\n"
        "出荷羽数がすべて空欄の場合は警告を記録します。\n"
        "出荷羽数合計が0の場合は警告を記録します。\n"
        "品種の明細がない場合は警告を記録します。\n"
        "警告がある場合、入力Excelと同じフォルダーに\n"
        "警告内容を記載したwarningファイルを作成します。\n"
        "step0002品種別ファイル名には「品種別」を含めます。\n"
        "品種別ファイル名には対象の年と月を含めます。\n"
        "品種別ファイル名の例：詳細データA_step0002_品種別_2026_07.tsv\n"
        "step0002総計ファイルも月ごとに作成します。\n"
        "step0002総計ファイル名には対象の年と月を含めます。\n"
        "品種別・総計の集計対象はそれぞれ同じ月の明細です。\n"
        "同名ファイルがある場合は上書きせず枝番を付けます。\n"
        "TSVはUTF-8（BOMなし）、タブ区切り、CRLF改行です。\n"
        "品種別TSVは追加出力であり、後続のグループ別処理には渡しません。\n"
        "品種別step0002から最新年を含む直近5年間の既存43項目を抽出し、\n"
        "品種別step0003の項目別TSVを出力します。\n"
        "品種別の年間集計系列はstep0002の1・2・3・5・6行目を使用します。\n"
        "ブロイラー合算行は使用せず、貞光分と外部購入を区別してstep0010まで引き継ぎます。\n"
        "それ以外の系列は従来どおり合算分を使用します。\n"
        "入雛羽数と出荷羽数を品種ごとに照合し、\n"
        "出荷羽数 ÷ 入雛羽数で廃棄控除前育成率を計算します。\n"
        "廃棄控除前育成率は100倍せず、小数点以下4桁で四捨五入します。\n"
        "入雛羽数が0の場合、廃棄控除前育成率は0です。\n"
        "出力名はstep0003_品種別_廃棄控除前育成率_年月.tsvです。\n"
        "品種別step0003の入雛羽数・出荷羽数・正味重量・平均体重、\n"
        "飼育日数平均、廃棄控除前育成率の6項目を品種ごとに結合し、\n"
        "品種別step0004として5区分を出力します。\n"
        "品種別step0004の列は年・月・品種No.・品種名に続けて、\n"
        "入雛羽数・出荷羽数・正味重量・平均体重・飼育日数平均・\n"
        "廃棄控除前育成率の順です。\n"
        "品種別step0004を品種ごと・年ごとに月順で連結し、\n"
        "step0005品種別年間TSVを出力します。\n"
        "step0005品種別年間TSVを行列入れ替えし、\n"
        "step0006品種別年間TSVを5区分出力します。\n"
        "5区分のstep0006の年・月を照合して項目行をまとめ、\n"
        "step0007品種別年間TSVを年ごとに出力します。\n"
        "step0006の最終月から12月までの不足月も列を補完します。\n"
        "品種別step0007年間TSVを読み込み、既存の\n"
        "廃棄控除前育成率の行名を育成率に変更してstep0008を作成します。\n"
        "計算値と行数は維持し、育成率行を追加しません。\n"
        "品種別step0008年間TSVを読み込み、5区分の月別データを保ったまま、\n"
        "各月の品種合計行を加えたstep0009品種別年間TSVを出力します。\n"
        "入雛羽数・出荷羽数・正味重量は合算し、平均体重と育成率は\n"
        "合算値から再計算します。飼育日数平均は出荷羽数で加重平均します。\n"
        "品種別step0009年間TSVの各品種と合計の各月データから、\n"
        "右端に年間合計列を追加したstep0010品種別年間TSVを作成します。\n"
        "羽数・重量は月合計、平均体重と育成率は年間合計から再計算し、\n"
        "飼育日数平均は出荷羽数による加重平均を出力します。\n"
        "step0013品種別年間TSVを読み込み、品種別シートの年間Excelを\n"
        "品種ごと・年ごとに出力します。\n"
        "step0013品種別Excelの各列幅を自動調整します。\n"
        "作成したExcelを同じフォルダに〇品種別実績集計_YYYY年.xlsx名でもコピーします。\n"
        "コピーした48行構成のExcelに控除廃棄版と同じ外枠・列区切り・横罫線を設定し、上書き保存します。\n"
        "対象年は元Excelのファイル名から取得し、罫線以外の内容・書式を保持します。\n"
        "同名時は_02などの連番を付け、元Excelは変更しません。\n"
        "通常年48行・2022年39行に対応し、稼動日数・前年実績・前年比を反映します。\n"
        "前年比は倍率を小数点以下4桁で表示し、通常版step0010年間Excelは作成しません。\n"
        "品種別step0011当年表に空行2行と前年の合計6項目・稼動日数の7行を追加し、\n"
        "step0012品種別年間TSVを出力します。\n"
        "当年step0012内の前年実績を使い、40・41行目に一日平均出荷羽数の前年比を表示した\n"
        "step0013品種別年間TSVを出力します。前年ファイルは再読込しません。\n"
        "前年比は倍率で小数点以下4桁に四捨五入し、計算不能は空欄、2022年はそのままコピーします。\n"
        "当年・前年の稼動日数の年間合計をO列に整数で設定します。\n"
        "前年欄は前年実績・前年の年表記とし、2022年は前年行と空行を追加しません。\n"
        "入雛羽数・処理羽数・処理重量・控除後平均体重・\n"
        "飼育日数平均・廃棄控除後育成率を品種ごとに結合し、\n"
        "控除廃棄版step0004を5区分出力します。\n"
        "控除廃棄版step0004を品種ごと・年ごとに月順で連結し、\n"
        "控除廃棄版step0005品種別年間TSVを5区分出力します。\n"
        "控除廃棄版step0005品種別年間TSVを行列入れ替えし、\n"
        "控除廃棄版step0006品種別年間TSVを5区分出力します。\n"
        "5区分の控除廃棄版step0006の年・月を照合して項目行をまとめ、\n"
        "控除廃棄版step0007品種別年間TSVを年ごとに出力します。\n"
        "最終月以降は12月までの不足月も列を補完します。\n"
        "控除廃棄版step0007品種別年間TSVを読み込み、既存の\n"
        "廃棄控除後育成率の行名を育成率に変更してstep0008を作成します。\n"
        "計算値と行数は維持し、育成率行を追加しません。\n"
        "控除廃棄版step0008品種別年間TSVを読み込み、5区分を保持して\n"
        "各月に合計行を追加したstep0009品種別年間TSVを年ごとに出力します。\n"
        "入雛羽数・処理羽数・処理重量は合算し、控除後平均体重は\n"
        "合計処理重量 ÷ 合計処理羽数で再計算します。\n"
        "飼育日数平均は処理羽数で加重平均し、育成率は\n"
        "合計処理羽数 ÷ 合計入雛羽数で再計算します。\n"
        "控除廃棄版step0009品種別年間TSVの各品種と合計について、\n"
        "羽数・重量は年間合計、控除後平均体重と育成率は年間値から再計算し、\n"
        "飼育日数平均は処理羽数による加重平均を追加した\n"
        "step0010品種別控除廃棄年間TSVを年ごとに出力します。\n"
        "控除廃棄版step0013品種別年間TSVを読み込み、\n"
        "品種別シートの年間Excelファイルを年ごとに出力します。\n"
        "控除廃棄版step0013品種別Excelの各列幅を自動調整します。\n"
        "作成したExcelを同じフォルダに〇品種別実績集計_控除廃棄_YYYY年.xlsx名でもコピーします。\n"
        "コピーした48行構成のExcelに見本と同じ外枠・列区切り・横罫線を設定し、上書き保存します。\n"
        "対象年は元Excelのファイル名から取得し、罫線以外の内容・書式を保持します。\n"
        "同名時は_02などの連番を付け、元Excelは変更しません。\n"
        "Excelは通常年48行・2022年39行に対応し、合計・稼動日数・前年実績・前年比を反映します。\n"
        "前年比は倍率を小数点以下4桁で表示し、step0010の年間Excelは作成しません。\n"
        "控除廃棄版品種別step0010と同年・同接頭辞の日別稼動判定TSVを読み、\n"
        "全品種共通の稼動日数行を追加した39行のstep0011を作成します。年間稼動日数は空欄です。\n"
        "step0011は同名時に連番で保存し、今回作成したファイルを後続処理へ渡します。\n"
        "当年・前年step0011を読み、空行2行と前年合計6項目・稼動日数の7行を追加した\n"
        "step0012品種別控除廃棄年間TSVを出力します。\n"
        "控除廃棄版step0012内の前年実績から一日平均処理羽数の前年比を計算し、\n"
        "40・41行目を前年比と補足表示に置き換えたstep0013品種別控除廃棄年間TSVを出力します。\n"
        "前年の別ファイルは読まず、倍率を小数点以下4桁に四捨五入し、計算不能は空欄とします。\n"
        "2022年のstep0013はstep0012をそのままコピーします。同名出力は上書きしません。\n"
        "飼育日数平均と稼動日数を残し、当年・前年の稼動日数の年間合計をO列に設定します。\n"
        "2022年は前年を読まず39行で出力します。同名出力は上書きしません。\n"
        "外部購入の日令はTSVに保持し、最終Excelの飼育日数平均だけ空欄にします。\n"
        "同じ対象月の品種別step0002のE列・F列・T列から、\n"
        "農家廃棄羽数・農家廃棄重量・日令も品種別step0003へ追加出力します。\n"
        "農家廃棄羽数はE列、農家廃棄重量はF列、日令はT列から抽出します。\n"
        "日令step0003ファイルを独立処理で複製し、\n"
        "飼育日数平均のファイル名でも同じ年月のTSVを出力します。\n"
        "複製ファイルの列見出しと内容は日令ファイルと同一です。\n"
        "各ファイルは年・月・品種No.・品種名・項目値の5列です。\n"
        "阿波尾鶏、赤どり、地養鳥、ブロイラー（貞光分）、ブロイラー(外部購入)の順で5行を出力します。\n"
        "同じ対象月の品種別step0003から廃棄15項目を合計し、\n"
        "全廃棄羽数・全廃棄重量・処理羽数・処理重量・控除後平均体重を、\n"
        "品種別step0003の項目別TSVとして追加出力します。\n"
        "全5品種区分の処理羽数は、出荷羽数から既存の全廃棄羽数と懸鳥前死鶏羽数・懸鳥前発育不良羽数を差し引きます。\n"
        "元ファイル名・年月・品種No.・品種名を照合し、入力の不足・重複・空欄・不正値はエラーにします。\n"
        "品種別自主廃棄重量はstep0001の明細ごとにマレック重量～予備４重量を合計し、0.72で割って小数第2位以下を切り捨ててから集計します。\n"
        "処理重量は出荷重量から懸鳥前重量2項目と換算後自主廃棄重量を差し引き、小数第1位に四捨五入します。\n"
        "修正後の羽数・重量を控除後平均体重と後続集計へ引き継ぎ、控除廃棄育成率の計算は維持します。\n"
        "ファイル名の控除後平均体重は、既存43項目の平均体重と区別します。\n"
        "全廃棄重量は合計後に小数点以下1桁で四捨五入します。\n"
        "廃棄控除後の平均体重は小数点以下8桁で四捨五入します。\n"
        "処理羽数が0の場合、廃棄控除後の平均体重は0です。\n"
        "既存step0003の平均体重は廃棄控除前、控除後平均体重は廃棄控除後の値です。\n"
        "同じ対象月の入雛羽数・処理羽数を品種ごとに照合し、\n"
        "処理羽数を入雛羽数で割った控除廃棄育成率を出力します。\n"
        "控除廃棄育成率は100倍せず、小数点以下4桁で四捨五入します。\n"
        "入雛羽数が0の場合、控除廃棄育成率は0です。\n"
        "出力名はstep0003_品種別_廃棄控除後育成率_年月.tsvです。\n"
        "グループ別・総計の後続処理は従来どおり実行します。\n"
        "最新年を含む直近5年間の\n"
        "指定43項目をグループ別・総計の項目別TSVへ抽出し、\n"
        "総計の入雛羽数と平均体重を横結合し、\n"
        "総計の出荷羽数と出荷重量を横結合し、\n"
        "総計の羽数15項目を項目別に直近5年間の月別表へまとめた\n"
        "step0004過去05年TSVを出力します。\n"
        "step0004の15ファイルを指定順に縦結合し、項目間に空白行を入れた\n"
        "step0005総計出荷廃棄羽数過去05年TSVも出力します。\n"
        "step0005総計出荷廃棄羽数過去05年TSVを読み込んだ\n"
        "同名のExcelファイルも出力します。\n"
        "作成したExcelを同じフォルダに〇出荷廃棄_羽数_過去05年.xlsx名でもコピーします。\n"
        "コピーしたExcelの総計羽数過去5年シートの15表（A～M列・各6行）に細い格子罫線を設定し、上書き保存します。\n"
        "同じExcelの「グラフ」シートに15項目・過去5年の折れ線グラフを2列×8段（最後の右側は空き）で作成します。\n"
        "タイトル・羽数の縦軸・1月～12月の横軸と年別凡例を表示し、作成対象年月より後の月・欠損月は線を描かず、実績0は描画します。\n"
        "再実行時は作成済みグラフを更新し、同名シートに別内容がある場合は保護してエラーにします。\n"
        "罫線以外の内容・書式を保持し、同名時は_02などの連番を付け、元Excelは変更しません。\n"
        "総計の廃棄羽数・重量各15項目を横結合し、\n"
        "各グループの廃棄羽数・廃棄重量各15項目を横結合したTSVも出力します。\n"
        "各グループの廃棄羽数15項目に小計列を追加した\n"
        "step0005グループ別TSVも出力します。\n"
        "各グループの廃棄重量15項目に小計列を追加した\n"
        "step0005グループ別TSVも出力します。\n"
        "step0005グループ別廃棄重量TSVの3グループを縦結合した\n"
        "step0006グループ別廃棄重量TSVを出力します。\n"
        "step0006グループ別廃棄重量TSVから全廃棄重量を抽出した\n"
        "step0007グループ別TSVを出力します。\n"
        "step0005グループ別廃棄羽数TSVの3グループを縦結合した\n"
        "step0006グループ別廃棄羽数TSVを出力します。\n"
        "step0006グループ別廃棄羽数TSVから全廃棄羽数を抽出した\n"
        "step0007グループ別TSVを出力します。\n"
        "step0004グループ別データにstep0007の全廃棄羽数・重量を加えた\n"
        "step0008グループ別控除廃棄TSVを出力します。\n"
        "グループ別step0003の懸鳥前死鶏羽数・懸鳥前発育不良羽数を照合し、step0008の11・12列に追加します。\n"
        "懸鳥前重量2項目と自主廃棄重量（0.72換算後）を末尾に追加し、step0008を15列にします。\n"
        "自主廃棄重量はstep0001の明細ごとにマレック重量～予備４重量を合計し、0.72で割って小数第2位以下を切り捨ててからグループ別に合計します。\n"
        "必要な入力の不足・重複・空欄はエラーとし、旧形式の10列・12列TSVはstep0008から再作成します。\n"
        "出力列の8列目は出荷重量です。\n"
        "契約・直営・フリーの処理羽数は、出荷羽数から全廃棄羽数と懸鳥前羽数2項目を差し引きます。\n"
        "処理重量は出荷重量から懸鳥前重量2項目と換算後自主廃棄重量を差し引き、小数第1位に四捨五入して出力します。\n"
        "平均体重と後続集計には四捨五入後の処理重量を使用します。\n"
        "処理羽数・処理重量と平均体重を計算した\n"
        "年・月を含むstep0009グループ別控除廃棄TSVを出力します。\n"
        "step0009契約・直営・フリー各グループの月別行を月順に連結した\n"
        "step0010グループ別控除廃棄年間TSVを出力します。\n"
        "step0010年間表を行列入れ替えしたstep0011グループ別控除廃棄TSVを出力します。\n"
        "step0011の契約・直営・フリーを年別にまとめたstep0012グループ別控除廃棄年間TSVを出力します。\n"
        "育成率を追加したstep0012からstep0013グループ別控除廃棄年間TSVを出力します。\n"
        "step0013グループ別控除廃棄年間TSVからExcelは作成しません。\n"
        "step0013の月別値から年次合計を計算し、\n"
        "平均体重・育成率も年間値から計算したstep0014グループ別控除廃棄年間TSVを出力します。\n"
        "全対象年のstep0014でO1は空欄、O2は「合計」とし、後続TSV・最終Excelへ引き継ぎます。\n"
        "O2が空欄の旧形式TSVはエラーとし、step0014からの再作成を案内します。\n"
        "step0014グループ別控除廃棄年間Excelは作成しません。\n"
        "step0014年間TSVとstep0001日別稼動判定TSVから稼動日数行を追加した\n"
        "step0015グループ別控除廃棄年間TSVを出力します。\n"
        "総計の懸鳥前羽数・重量各2項目を横結合し、\n"
        "懸鳥前・廃棄データ別の羽数・重量TSVに小計列を追加します。\n"
        "入雛・懸鳥前・廃棄羽数データを横結合した\n"
        "廃棄データ一覧表の上段を出力します。\n"
        "出荷重量・懸鳥前・廃棄重量データを横結合した\n"
        "廃棄データ一覧表の下段を出力します。\n"
        "上段・下段それぞれの小計2列を合算し、\n"
        "step0007廃棄データ一覧表に合計列を追加します。\n"
        "step0007の上段・下段を1つにまとめた\n"
        "step0008総計廃棄データ一覧表を出力します。\n"
        "step0008総計一覧表からグループ列を除き、\n"
        "羽数・重量の区分列を追加したstep0009を出力します。\n"
        "step0009の月別一覧表を年月の新しい順にまとめた\n"
        "step0010総計廃棄データ一覧表を出力し、\n"
        "1行目のH列・X列・Y列の見出しを空白にします。\n"
        "step0010表の右端に前年同月比(羽数)列を追加した\n"
        "step0011総計廃棄データ一覧表を出力します。\n"
        "前年同月の羽数合計との比率をZ列に追加した\n"
        "step0012総計廃棄データ一覧表を出力します。\n"
        "step0012総計廃棄データ一覧表TSVからExcelファイルを出力し、\n"
        "羽数行のD列（平均体重）は小数点以下2桁で表示します。\n"
        "各列の見出しとデータに合わせて列幅を自動調整します。\n"
        "ExcelのA列・B列から最新年月を取得し、\n"
        "〇廃棄データ月別一覧表～YYYY年MM月.xlsxの名前でもコピーを出力します。\n"
        "年月付きExcelを読み込み、外枠・列区切り・月別の罫線を設定して上書き保存します。\n"
        "数値・数式・フォント・表示形式・列幅は変更しません。\n"
        "年月付き一覧表に「グラフ」シートを追加し、羽数8項目の過去5年折れ線グラフを2列×4段で配置します。\n"
        "各グラフの項目タイトル、縦軸の数値と「羽数」、横軸の1月～12月を表示し、文字と凡例が重ならないよう配置します。\n"
        "未到来月・欠損月は線を描かず、実績0は描画します。既存の「総計」シートは維持します。\n"
        "再実行は作成済みグラフを更新し、同名シートに別の内容がある場合はエラーにします。\n"
        "主要4項目をグループ別TSVへ結合し、\n"
        "年・グループ別の年間TSVへ連結し、\n"
        "年間TSVの行と列を入れ替え、\n"
        "契約・直営・フリーを年別に連結します。\n"
        "step0008グループ別年間TSVからExcelは作成しません。\n"
        "育成率追加後の年間TSVへ合計を追加します。\n"
        "step0009グループ別年間TSVからExcelは作成しません。\n"
        "年間TSVの右端に年間合計列を追加します。\n"
        "年間合計列付きTSVからExcelを作成します。\n"
        "step0010グループ別年間TSVとstep0001日別稼動判定TSVから稼動日数行を追加した\n"
        "step0011グループ別年間TSVを出力します。\n"
        "step0011当年表に空行2行と前年の合計6行を加え、step0012グループ別年間TSVを出力します。\n"
        "前年欄は前年実績・前年の年表記とし、2022年は前年行を追加しません。\n"
        "step0017グループ別控除廃棄年間TSVを読み込み、同じファイル名のExcelを作成します。\n"
        "作成したExcelを同じフォルダに〇グループ別集計_控除廃棄_YYYY年.xlsx名でもコピーします。\n"
        "コピーした31行構成のExcelに見本と同じ外枠・縦線・横線を設定し、上書き保存します。\n"
        "対象年は元Excelのファイル名から取得し、罫線以外の内容・書式を保持します。\n"
        "同名時は_02などの連番を付け、元Excelは変更しません。\n"
        "各列の見出しとデータに合わせて列幅を自動調整します。\n"
        "step0012の当年・前年の稼動日数合計をO列に整数で出力します。\n"
        "step0012の24・25行目に一日平均出荷羽数の前年比を追加したstep0013を出力します。\n"
        "前年比は小数点以下4桁で四捨五入し、計算不能は空欄、2022年はそのままコピーします。\n"
        "step0013グループ別年間TSVを読み込み、同じstep0013名のExcelを作成します。\n"
        "step0013グループ別Excelの各列幅を自動調整し、step0010グループ別Excelは作成しません。\n"
        "step0013グループ別Excelを年別の〇グループ別集計_YYYY年.xlsx名でも出力します。\n"
        "対象年は元Excelのファイル名から取得し、同じフォルダへコピーします。\n"
        "コピーした31行構成のExcelに見本と同じ外枠・縦線・横線を設定して上書き保存し、罫線以外の内容・書式を保持します。\n"
        "同名ファイルがある場合は_02などの連番を付け、元Excelは変更しません。\n"
        "step0015控除廃棄当年表にも空行2行と前年の合計6行を加え、step0016を出力します。\n"
        "項目は入雛羽数・処理羽数・処理重量・平均体重・育成率・稼動日数です。\n"
        "前年欄は前年実績・前年の年表記とし、2022年は前年行を追加しません。\n"
        "step0016の当年・前年の稼動日数合計をO列に整数で出力します。\n"
        "step0016の24・25行目に一日平均処理羽数の前年比を追加したstep0017を出力します。\n"
        "前年比は小数点以下4桁で四捨五入し、計算不能は空欄、2022年はそのままコピーします。\n"
        "step0001のO列「処理日」から年別の日別稼動判定TSVを追加出力します。\n"
        "1行目は対象年（例：2026年）と半角数字の1～12月、2行目は稼動日数、3～33行目は1～31日の0／1です。\n"
        "未提供月は空欄、存在しない日は0です。判定TSVはUTF-8 BOM付きです。\n"
        "その他のシートは処理しません。\n\n"
        "1回につき1ファイルだけ処理できます。\n"
        "出力TSVは入力Excelと同じフォルダに作成します。\n\n"
        "対応形式: .xlsx / .xlsm"
    )
    win32gui.DrawText(
        iDeviceContextHandle,
        pszInstructionText,
        -1,
        objClientRect,
        win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK,
    )
    win32gui.EndPaint(iWindowHandle, objPaintStruct)


def window_proc(iWindowHandle: int, iMessage: int, iWparam: int, iLparam: int) -> int:
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0
    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        try:
            iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)
            arrInputPaths: List[str] = [
                os.path.abspath(win32api.DragQueryFile(iDropHandle, iIndex))
                for iIndex in range(iFileCount)
            ]
        finally:
            win32api.DragFinish(iDropHandle)

        if not arrInputPaths:
            show_error_message_box("Error: no Excel file was dropped.")
            return 0
        if len(arrInputPaths) != 1:
            show_error_message_box(
                "Excelファイルは1回につき1ファイルだけ\n"
                "ドラッグ＆ドロップしてください。\n\n"
                "ドロップされたファイル数: " + str(len(arrInputPaths))
            )
            return 0
        run_excel_to_tsv(arrInputPaths[0])
        return 0
    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0
    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def main() -> None:
    pszWindowClassName: str = "GroupShipmentSummaryDndWindowClass"
    iInstanceHandle: int = win32api.GetModuleHandle(None)
    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1
    win32gui.RegisterClass(objWndClass)

    iWindowHandle: int = win32gui.CreateWindowEx(
        win32con.WS_EX_ACCEPTFILES,
        pszWindowClassName,
        WINDOW_TITLE,
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        520,
        320,
        0,
        0,
        iInstanceHandle,
        None,
    )
    win32gui.ShowWindow(iWindowHandle, win32con.SW_SHOWNORMAL)
    win32gui.UpdateWindow(iWindowHandle)
    win32gui.SetWindowPos(
        iWindowHandle,
        win32con.HWND_TOPMOST,
        0,
        0,
        0,
        0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
    )
    win32gui.PumpMessages()


if __name__ == "__main__":
    main()
