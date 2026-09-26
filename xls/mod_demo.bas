Attribute VB_Name = "mod_demo"
'=====================================================================
' mod_demo - 线性回归完整演示（造数据 → 训练 → 预测 → 评估）
'=====================================================================
' 用途：一条龙展示 mod_ml 的轻量机器学习能力，跑一遍即可看懂全流程：
'   1) Ml_LinearReg_Sample   造带噪声的模拟数据（真系数已知，seed 固定可复现）
'   2) Ml_LinearReg_Train    最小二乘训练出模型系数（与真系数对比看拟合效果）
'   3) Ml_LinearReg_Predict  对新样本预测（与解析真值对比）
'   4) Ml_LinearReg_RSquared 输出拟合优度 R2（0~1，越接近 1 拟合越好）
'
' 用法：把库导入任意 .xlsm（src 纯层 + xls 绑定层）后运行宏：
'       Demo_Ml_LinearReg
'   自动在活动工作簿新建"线性回归演示"工作表，输出：数据表 / 模型系数 / 预测对比，
'   同时在立即窗口（Ctrl+G）打印全部关键数值。
'
' 依赖：mod_ml（纯层算法）、mod_range（批量写单元格）。
'=====================================================================
Option Explicit

' 主演示入口：造数据 → 训练 → 预测 → 评估 → 写回工作表。
Public Sub Demo_Ml_LinearReg()
    ' ---- 0. 演示配置：真系数 y = 2 + 3·x1 - 1·x2 + 噪声 ----
    Const N_SAMPLES As Long = 60
    Const N_FEATURES As Long = 2
    Const NOISE_AMP As Double = 0.8
    Const SEED As Long = 7
    Dim true_coefs As Variant
    true_coefs = Array(2#, 3#, -1#)   ' [截距 b0, 权重 w1, 权重 w2]

    ' ---- 1. 造数据：60 行 x 3 列（特征1 特征2 | 标签 y） ----
    Dim smp As Variant
    smp = mod_ml.Ml_LinearReg_Sample(N_SAMPLES, N_FEATURES, NOISE_AMP, true_coefs, SEED)
    Debug.Print "=== 线性回归演示 ==="
    Debug.Print "造数据: " & N_SAMPLES & " 样本 x " & N_FEATURES & " 特征, 噪声 ±" & NOISE_AMP & ", seed=" & SEED

    ' ---- 2. 训练：最小二乘拟合，返回 0 基系数数组 [截距, 权重...] ----
    Dim coef As Variant
    coef = mod_ml.Ml_LinearReg_Train(smp)
    Debug.Print "训练系数: 截距=" & Round(coef(0), 3) & _
                "  w1=" & Round(coef(1), 3) & _
                "  w2=" & Round(coef(2), 3) & _
                "  (真值: 2, 3, -1)"

    ' ---- 3. 预测：3 个新样本（未参与训练），与解析真值对比 ----
    Dim xs1 As Variant, xs2 As Variant, xs3 As Variant
    xs1 = Array(1.5, 2.5): xs2 = Array(4#, 3#): xs3 = Array(7#, 0.5)
    Dim y_truth1 As Double, y_truth2 As Double, y_truth3 As Double
    Dim y_hat1 As Double, y_hat2 As Double, y_hat3 As Double
    y_truth1 = 2# + 3# * 1.5 - 1# * 2.5      ' 解析值 = 4.0
    y_truth2 = 2# + 3# * 4# - 1# * 3#        ' 解析值 = 11.0
    y_truth3 = 2# + 3# * 7# - 1# * 0.5       ' 解析值 = 22.5
    y_hat1 = mod_ml.Ml_LinearReg_Predict(coef, xs1)
    y_hat2 = mod_ml.Ml_LinearReg_Predict(coef, xs2)
    y_hat3 = mod_ml.Ml_LinearReg_Predict(coef, xs3)
    Debug.Print "预测对比: 样本(1.5,2.5) 解析=" & Round(y_truth1, 3) & " 模型=" & Round(y_hat1, 3)
    Debug.Print "          样本(4,3)    解析=" & Round(y_truth2, 3) & " 模型=" & Round(y_hat2, 3)
    Debug.Print "          样本(7,0.5)  解析=" & Round(y_truth3, 3) & " 模型=" & Round(y_hat3, 3)

    ' ---- 4. 评估：拟合优度 R2（用训练数据，低噪声下应接近 1） ----
    Dim r2 As Double
    r2 = mod_ml.Ml_LinearReg_RSquared(smp, coef)
    Debug.Print "拟合优度 R2 = " & Round(r2, 4)

    ' ---- 5. 写回工作表"线性回归演示" ----
    Call write_demo_sheet(smp, coef, y_truth1, y_hat1, y_truth2, y_hat2, y_truth3, y_hat3, r2)
    Debug.Print "=== 完成：结果已写入工作表『线性回归演示』 ==="
End Sub

' 私有辅助：在活动工作簿建/取"线性回归演示"表并输出数据表、系数、预测对比。
Private Sub write_demo_sheet(ByRef smp As Variant, ByRef coef As Variant, _
                             ByVal yt1 As Double, ByVal yh1 As Double, _
                             ByVal yt2 As Double, ByVal yh2 As Double, _
                             ByVal yt3 As Double, ByVal yh3 As Double, _
                             ByVal r2 As Double)
    Dim wb As Workbook, ws As Worksheet
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets("线性回归演示")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "线性回归演示"
    End If

    ' ---- 数据表：表头 + 全部样本，一次批量写入 ----
    Dim n As Long, f As Long, i As Long, j As Long
    n = UBound(smp, 1) - LBound(smp, 1) + 1
    f = UBound(smp, 2) - LBound(smp, 2) + 1
    Dim tbl() As Variant
    ReDim tbl(1 To n + 1, 1 To f)
    For j = 1 To f - 1
        tbl(1, j) = "特征" & j
    Next j
    tbl(1, f) = "y(标签)"
    For i = 1 To n
        For j = 1 To f
            tbl(i + 1, j) = smp(i, j)
        Next j
    Next i
    Call mod_range.Range_WriteValues(ws.Range("A1"), tbl)

    ' ---- 模型系数区：截距 / 权重 / 拟合优度 ----
    Dim base_row As Long
    base_row = n + 3
    ws.Range("A" & base_row).Value = "【模型系数】 真值: 截距=2, w1=3, w2=-1"
    ws.Range("A" & base_row + 1).Value = "截距 b0"
    ws.Range("B" & base_row + 1).Value = Round(coef(0), 4)
    ws.Range("A" & base_row + 2).Value = "权重 w1"
    ws.Range("B" & base_row + 2).Value = Round(coef(1), 4)
    ws.Range("A" & base_row + 3).Value = "权重 w2"
    ws.Range("B" & base_row + 3).Value = Round(coef(2), 4)
    ws.Range("A" & base_row + 4).Value = "拟合优度 R2"
    ws.Range("B" & base_row + 4).Value = Round(r2, 4)

    ' ---- 预测对比表：新样本 / 解析真值 / 模型预测 ----
    Dim pred_base As Long
    pred_base = base_row + 6
    ws.Range("A" & pred_base - 1).Value = "【预测对比】 新样本（未参与训练）"
    Dim ptab() As Variant
    ReDim ptab(1 To 4, 1 To 4)
    ptab(1, 1) = "x1": ptab(1, 2) = "x2": ptab(1, 3) = "解析 y": ptab(1, 4) = "模型预测"
    ptab(2, 1) = 1.5: ptab(2, 2) = 2.5: ptab(2, 3) = yt1: ptab(2, 4) = yh1
    ptab(3, 1) = 4#: ptab(3, 2) = 3#: ptab(3, 3) = yt2: ptab(3, 4) = yh2
    ptab(4, 1) = 7#: ptab(4, 2) = 0.5: ptab(4, 3) = yt3: ptab(4, 4) = yh3
    Call mod_range.Range_WriteValues(ws.Range("A" & pred_base), ptab)

    ' ---- 轻量美化：表头加粗、按内容调列宽 ----
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A" & pred_base & ":D" & pred_base).Font.Bold = True
    ws.Range("A" & base_row).Font.Bold = True
    ws.Range("A" & pred_base - 1).Font.Bold = True
    ws.Columns("A:D").AutoFit
    ws.Activate
End Sub
