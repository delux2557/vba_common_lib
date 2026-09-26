Attribute VB_Name = "mod_ml"
'=====================================================================
' mod_ml - 轻量机器学习（通用库 · v1.2）
'=====================================================================
' 命名体系：公共 API 使用 PascalCase + 域名前缀（Ml_*），内部过程与局部变量使用 snake_case。
'
' 职责：在 Excel 内做轻量数据分析的经典算法。纯计算、无宿主依赖，归入 src/ 纯层。
'       定位是"轻量、可解释、中小数据"：数据进出用 Range_Read/WriteValues 衔接。
'
' 约定：
'   · 数据输入为二维数组（1 基，来自 Range_ReadValues 或 Ml_LinearReg_Sample），
'     默认"末列为标签、其余列为特征"；label_col 可选参数可覆盖（1 基列号）。
'   · 模型产物一律为普通数组（如系数数组），可写回单元格展示，也可再传给预测函数。
'   · 训练与预测分离：Train 只做一次，Predict 可反复接新样本。
'
' 目录 / Catalog（新函数加入时在此登记）
'   Ml_LinearReg_Sample(n[, features, noise, coefs, seed]) As Variant  生成线性样本数据集
'   Ml_LinearReg_Train(data[, label_col])                 As Variant  最小二乘训练,返回系数数组[截距,权重...]
'   Ml_LinearReg_Predict(model, sample)                   As Double    对新样本预测
'   Ml_LinearReg_RSquared(data, model[, label_col])       As Double    评估拟合优度 R2(0~1)
'=====================================================================
Option Explicit

' 生成带噪声的线性回归样本数据集：y = b0 + Σ w_j·x_j + 噪声。
' 参数 n 样本数、features 特征数、noise 噪声幅度(±noise 均匀噪声)。
' coefs 可选：真系数数组 [b0, w1, ...]，长度须为 features+1——训练后可与它对比检验效果；
'       不传则真系数随机取 -3~3。
' seed 可选：>0 时固定随机序列（测试/演示可复现），=0 默认每次随机。
' 返回 1 基二维数组：前 features 列是特征(0~10 均匀)，末列是标签 y。
Public Function Ml_LinearReg_Sample(ByVal n As Long, Optional ByVal features As Long = 1, _
                                    Optional ByVal noise As Double = 1#, _
                                    Optional ByVal coefs As Variant, _
                                    Optional ByVal seed As Long = 0) As Variant
    Dim b0 As Double, w() As Double, j As Long
    If n < 2 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Sample: 样本数须 ≥ 2"
    If features < 1 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Sample: 特征数须 ≥ 1"
    If noise < 0 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Sample: 噪声幅度须 ≥ 0"
    If seed > 0 Then Randomize seed Else Randomize
    ' 真系数：调用方传入（可复现对比）或随机取 -3~3
    ReDim w(1 To features)
    If Not IsMissing(coefs) Then
        If Not IsArray(coefs) Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Sample: coefs 须为数组"
        If UBound(coefs) - LBound(coefs) + 1 <> features + 1 Then _
            Err.Raise 45000, "mod_ml", "Ml_LinearReg_Sample: coefs 长度须为 features+1（截距+权重）"
        b0 = CDbl(coefs(LBound(coefs)))
        For j = 1 To features
            w(j) = CDbl(coefs(LBound(coefs) + j))
        Next j
    Else
        b0 = rnd_signed
        For j = 1 To features
            w(j) = rnd_signed
        Next j
    End If
    ' 特征 ~ U[0,10]；y = b0 + Σ w_j·x_j + 噪声(±noise)
    Dim out() As Variant
    ReDim out(1 To n, 1 To features + 1)
    Dim i As Long, y As Double
    For i = 1 To n
        y = b0
        For j = 1 To features
            out(i, j) = 10 * Rnd
            y = y + w(j) * out(i, j)
        Next j
        y = y + (Rnd - 0.5) * 2 * noise
        out(i, features + 1) = y
    Next i
    Ml_LinearReg_Sample = out
End Function

' 最小二乘训练线性回归模型（正规方程 + 高斯消元，支持多特征）。
' data 为二维数组，默认末列是标签；label_col 可覆盖（1 基列号）。
' 样本数须 ≥ 特征数+1（欠定时无唯一解，抛错 #45000）。
' 返回 0 基系数数组：index 0 = 截距 b0，index 1.. = 各特征权重 w1..wf。
Public Function Ml_LinearReg_Train(ByRef data As Variant, Optional ByVal label_col As Long = 0) As Variant
    Dim n As Long, f As Long, label As Long, i As Long, c As Long, idx As Long
    If Not IsArray(data) Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Train: data 须为二维数组"
    If num_dims(data) <> 2 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Train: data 须为二维数组"
    n = UBound(data, 1) - LBound(data, 1) + 1
    f = UBound(data, 2) - LBound(data, 2) + 1
    If f < 2 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Train: 至少需要一列特征"
    If label_col <= 0 Then label_col = f
    If label_col < 1 Or label_col > f Then _
        Err.Raise 45000, "mod_ml", "Ml_LinearReg_Train: label_col 越界"
    Dim nfeat As Long
    nfeat = f - 1
    If n < nfeat + 1 Then _
        Err.Raise 45000, "mod_ml", "Ml_LinearReg_Train: 样本数须 ≥ 特征数+1（当前欠定）"
    ' 特征列绝对索引表（跳过标签列）
    Dim feat() As Long
    ReDim feat(1 To nfeat)
    idx = 0
    For c = 1 To f
        If c <> label_col Then
            idx = idx + 1
            feat(idx) = c
        End If
    Next c
    ' 正规方程 XtX·β = XtY（首列为截距列 1）
    Dim p As Long, r As Long, j As Long
    p = nfeat + 1
    Dim xtx() As Double
    ReDim xtx(1 To p, 1 To p)
    Dim xty() As Double
    ReDim xty(1 To p)
    Dim xr As Double, xc As Double, yv As Double
    For i = LBound(data, 1) To UBound(data, 1)
        yv = CDbl(data(i, label_col))
        For r = 1 To p
            If r = 1 Then xr = 1 Else xr = CDbl(data(i, feat(r - 1)))
            For c = r To p
                If c = 1 Then xc = 1 Else xc = CDbl(data(i, feat(c - 1)))
                xtx(r, c) = xtx(r, c) + xr * xc
                If c > r Then xtx(c, r) = xtx(c, r) + xr * xc
            Next c
            xty(r) = xty(r) + xr * yv
        Next r
    Next i
    ' 组装增广矩阵 [XtX | XtY] 并高斯消元求解
    Dim aug() As Double
    ReDim aug(1 To p, 1 To p + 1)
    For r = 1 To p
        For c = 1 To p
            aug(r, c) = xtx(r, c)
        Next c
        aug(r, p + 1) = xty(r)
    Next r
    Ml_LinearReg_Train = solve_linear(aug)
End Function

' 用训练好的系数模型对新样本预测。
' model 为 Ml_LinearReg_Train 返回的 0 基系数数组。
' sample 为一维数组（各特征值，长度 = 权重数）；单特征模型也接受标量。
' 返回预测值 = b0 + Σ w_j·x_j。
Public Function Ml_LinearReg_Predict(ByRef model As Variant, ByRef sample As Variant) As Double
    Dim m As Long, i As Long
    If Not IsArray(model) Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Predict: model 须为系数数组"
    m = UBound(model) - LBound(model) + 1
    If m < 2 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_Predict: model 至少含截距与一个权重"
    Dim acc As Double
    acc = model(LBound(model))
    If IsArray(sample) Then
        If UBound(sample) - LBound(sample) + 1 <> m - 1 Then _
            Err.Raise 45000, "mod_ml", "Ml_LinearReg_Predict: 样本特征数与模型权重数不符"
        For i = LBound(sample) To UBound(sample)
            acc = acc + model(LBound(model) + (i - LBound(sample)) + 1) * CDbl(sample(i))
        Next i
    Else
        If m <> 2 Then _
            Err.Raise 45000, "mod_ml", "Ml_LinearReg_Predict: 多特征模型须传一维数组样本"
        acc = acc + model(LBound(model) + 1) * CDbl(sample)
    End If
    Ml_LinearReg_Predict = acc
End Function

' 评估模型拟合优度 R2 = 1 - SS_res / SS_tot（0~1，越接近 1 拟合越好）。
' data 二维数组、model 系数数组，label_col 语义同 Train。
' 标签无变异（SS_tot=0）时 R2 无意义，抛错 #45000。
Public Function Ml_LinearReg_RSquared(ByRef data As Variant, ByRef model As Variant, _
                                      Optional ByVal label_col As Long = 0) As Double
    Dim n As Long, f As Long, label As Long, i As Long, c As Long, k As Long
    If Not IsArray(data) Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_RSquared: data 须为二维数组"
    If num_dims(data) <> 2 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_RSquared: data 须为二维数组"
    If Not IsArray(model) Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_RSquared: model 须为系数数组"
    n = UBound(data, 1) - LBound(data, 1) + 1
    f = UBound(data, 2) - LBound(data, 2) + 1
    If f < 2 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_RSquared: 至少需要一列特征"
    If label_col <= 0 Then label_col = f
    If label_col < 1 Or label_col > f Then _
        Err.Raise 45000, "mod_ml", "Ml_LinearReg_RSquared: label_col 越界"
    ' 标签均值（作为"只用均值预测"的基线）
    Dim sum_y As Double
    For i = LBound(data, 1) To UBound(data, 1)
        sum_y = sum_y + CDbl(data(i, label_col))
    Next i
    Dim mean_y As Double
    mean_y = sum_y / n
    Dim ss_res As Double, ss_tot As Double
    Dim samp() As Double
    ReDim samp(1 To f - 1)
    Dim yhat As Double
    For i = LBound(data, 1) To UBound(data, 1)
        k = 0
        For c = 1 To f
            If c <> label_col Then
                k = k + 1
                samp(k) = CDbl(data(i, c))
            End If
        Next c
        yhat = Ml_LinearReg_Predict(model, samp)
        ss_res = ss_res + (CDbl(data(i, label_col)) - yhat) ^ 2
        ss_tot = ss_tot + (CDbl(data(i, label_col)) - mean_y) ^ 2
    Next i
    If ss_tot = 0 Then Err.Raise 45000, "mod_ml", "Ml_LinearReg_RSquared: 标签无变异，R2 无意义"
    Ml_LinearReg_RSquared = 1 - ss_res / ss_tot
End Function

' 私有辅助：探测数组维度（1 基二维数组返回 2；标量/一维返回 0/1）。
Private Function num_dims(ByRef arr As Variant) As Long
    Dim d As Long, ub As Long
    On Error Resume Next
    Do
        d = d + 1
        Err.Clear
        ub = UBound(arr, d)
        If Err.Number <> 0 Then Exit Do
    Loop
    num_dims = d - 1
End Function

' 私有辅助：高斯消元（列主元）解线性方程组，增广矩阵 aug 为 1 基 (1..m, 1..m+1)。
' 返回 0 基解向量 x(0..m-1)。矩阵奇异时抛错 #45000。
Private Function solve_linear(ByRef aug As Variant) As Variant
    Dim m As Long, i As Long, j As Long, k As Long
    Dim p_row As Long, t As Double
    m = UBound(aug, 1)
    ' 前向消元：逐列选主元（取绝对值最大行，降低数值误差）
    For k = 1 To m
        p_row = k
        For i = k + 1 To m
            If Abs(aug(i, k)) > Abs(aug(p_row, k)) Then p_row = i
        Next i
        If Abs(aug(p_row, k)) < 1E-12 Then _
            Err.Raise 45000, "mod_ml", "solve_linear: 矩阵奇异（特征可能存在共线），无法求解"
        If p_row <> k Then
            For j = k To m + 1
                t = aug(k, j): aug(k, j) = aug(p_row, j): aug(p_row, j) = t
            Next j
        End If
        For i = k + 1 To m
            t = aug(i, k) / aug(k, k)
            For j = k To m + 1
                aug(i, j) = aug(i, j) - t * aug(k, j)
            Next j
        Next i
    Next k
    ' 回代求解
    Dim x() As Double
    ReDim x(0 To m - 1)
    For i = m To 1 Step -1
        x(i - 1) = aug(i, m + 1)
        For j = i + 1 To m
            x(i - 1) = x(i - 1) - aug(i, j) * x(j - 1)
        Next j
        x(i - 1) = x(i - 1) / aug(i, i)
    Next i
    solve_linear = x
End Function

' 私有辅助：均匀随机数 -3 ~ 3（用于随机真系数）。
Private Function rnd_signed() As Double
    rnd_signed = (Rnd - 0.5) * 6
End Function
