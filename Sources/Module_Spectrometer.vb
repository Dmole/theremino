
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices

Module Spectrometer

    Private MeterBar1 As MeterBar = New MeterBar(Form_Main.PictureBox3, 4)
    'm_fillbrush = CreateFillBrush(pbox, 0, 4)  angle=0  blend=4 red

    Friend Sub InitPictureboxImage(ByVal pbox As PictureBox)
        With pbox
            If .ClientSize.Height < 1 Then Return
            If .Image Is Nothing OrElse .Image.Width <> .ClientRectangle.Width OrElse .Image.Height <> .ClientRectangle.Height Then
                .Image = New Bitmap(.ClientRectangle.Width, .ClientRectangle.Height, Imaging.PixelFormat.Format24bppRgb)
            End If
        End With
    End Sub

    Friend TrimPoint1 As Single = 436
    Friend TrimPoint2 As Single = 546

    Friend ReferenceScale As Boolean
    Friend CursorInside As Boolean
    Friend MaxNanometers_OldValue As Int32

    Private ShowDips As Boolean
    Private ShowPeaks As Boolean
    Private UseColors As Boolean
    Private TrimScale As Boolean

    Private SpecV(-1) As Single '增加一个数组，存储原始的像素灰度值
    Private SpecArray(-1) As Single 'A zero-length array is declared with a dimension of -1.光谱阵列，Spec - spectrum，这里是积分之后数值
    Private SpecArrayFiltered(-1) As Single '过滤后的数值
    Private MaxValue As Single
    Private MaxValueX As Int32

    '我增加的变量，用来定义光谱能显示的最大值，每个px的三个颜色通道的和最大是765，
    '但增加一些积分时间或滤波系数后，可能会大一些
    Private MaxV As Single = 800.0

    Private Filter As Int32
    Private SpeedUP As Int32
    Private SpeedDW As Int32
    Private Flip As Boolean

    '下面这些参数都是加权值
    Private KFilter As Single
    Private KSpeedUP As Single
    Private KSpeedDW As Single
    Private Kred As Single
    Private Kgreen As Single
    Private Kblue As Single

    '这些是橘色选择框的尺寸和位置
    Private SrcY0 As Int32
    Private SrcDY As Int32
    Private SrcX0 As Int32
    Private SrcDX As Int32

    Private SrcImage As Image     'source 来源
    Private SrcBmp As Bitmap
    Private SrcW As Int32
    Private SrcH As Int32
    Private SrcPen As Pen

    Private DestPbox As PictureBox = Form_Main.PBox_Spectrum
    Private DestW As Int32
    Private DestH As Int32
    Private DestRight As Int32
    Private DestBottom As Int32
    Private ScalePen1 As Pen = New Pen(Color.FromArgb(140, 140, 140))
    Private ScalePen2 As Pen = New Pen(Color.FromArgb(200, 200, 200))
    Private ScalePen3 As Pen = New Pen(Color.FromArgb(220, 220, 220))
    Private ScaleFont As Font = New Font("Arial", 8)

    Friend NanometersMin As Single = 270
    Friend NanometersMax As Single = 1200
    Private NanometersDelta As Single  'Delta（Δ）在微积分中，Δ通常用来表示增量

    Private NmStart As Single
    Private NmEnd As Single
    Private NmStartDiv As Int32
    Private NmCoeff As Single 'Coeff，系数

    Private gfx As Graphics
    Private kx As Single
    Private ky As Single
    Private Pen_Graph As Pen = New Pen(Color.Black)


    Friend Sub Spectrometer_SetSourceParams()
        If SrcImage Is Nothing Then Return
        ' ---------------------------------------------------------------------
        SrcW = SrcImage.Width
        SrcH = SrcImage.Height

        '获取Filp的状态
        Flip = Form_Main.chk_Flip.Checked

        'CInt(n),将文本框输入的字符转换成Int类型；
        Dim StartY As Int32 = Form_Main.txt_StartY.NumericValueInteger
        Dim SizeY As Int32 = Form_Main.txt_SizeY.NumericValueInteger
        Dim StartX As Int32 = Form_Main.txt_StartX.NumericValueInteger
        Dim EndX As Int32 = Form_Main.txt_EndX.NumericValueInteger

        '获取橘色选择框的位置和尺寸
        SrcX0 = (SrcW * StartX) \ 1000   'StartX-EndX范围是1000，所以相当于换算成实际的像素尺寸
        SrcDX = SrcW - SrcX0 + (SrcW * (EndX - 1000)) \ 1000  '得出获取相机图像的像素长度，也就是橘色选择框的长度
        SrcDY = (SrcH * SizeY) \ 100    '橘色选择框的高度
        SrcY0 = SrcH - (SrcH * StartY \ 100) - SrcDY
        ' ---------------------------------------------------------------------
        If SrcX0 + SrcDX > SrcImage.Width Then SrcDX = SrcW - SrcX0
        If SrcDX <= 0 Then SrcX0 += (SrcDX - 1) : SrcDX = 1  '顺序执行以冒号分隔的语句
        If SrcX0 < 0 Then SrcX0 = 0
        If SrcY0 + SrcDY > SrcH Then SrcDY = SrcH - SrcY0
        If SrcDY <= 0 Then SrcY0 += (SrcDY - 1) : SrcDY = 1
        If SrcY0 < 0 Then SrcY0 = 0
        ' ---------------------------------------------------------------------
        ReDim SpecV(SrcDX - 1)
        ReDim SpecArray(SrcDX - 1)  '重新定义数组的长度
        ReDim SpecArrayFiltered(SrcDX - 1)

        Kred = 1.0F / SrcDY     '前值是0.5F，去拜尔后改成1.0F
        Kgreen = 1.0F / SrcDY   '前值是0.45F，去拜尔后改成1.0F
        Kblue = 1.0F / SrcDY    '前值是0.5F，去拜尔后改成1.0F
        SrcPen = New Pen(Color.FromArgb(200, 120, 0), SrcH \ 100) '绘制橘色选择框
        ' ---------------------------------------------------------------------
        Spectrometer_SetScaleTrimParams()
        Spectrometer_ResetReference() '取消参考光谱，Form_Main.btn_Reference.Checked = False
    End Sub

    Friend Sub Spectrometer_SetRunningModeParams()
        ShowDips = Form_Main.btn_Dips.Checked()
        ShowPeaks = Form_Main.btn_Peaks.Checked()
        UseColors = Form_Main.btn_Colors.Checked()
        TrimScale = Form_Main.btn_TrimScale.Checked()
        Filter = Form_Main.txt_Filter.NumericValueInteger
        KFilter = (100 - Filter) / 100.0F + 0.1F
        SpeedUP = Form_Main.txt_RisingSpeed.NumericValueInteger
        KSpeedUP = SpeedUP / 100.0F
        SpeedDW = Form_Main.txt_FallingSpeed.NumericValueInteger
        KSpeedDW = SpeedDW / 100.0F
    End Sub

    ' Version 2.9 - CHANGED from 2000 to 4000
    Friend Sub Spectrometer_SetScaleTrimParams()
        If NanometersMin < 50 Then NanometersMin = 50
        If NanometersMax > 4000 Then NanometersMax = 4000
        If NanometersMax < NanometersMin + 20 Then NanometersMax = NanometersMin + 20
        NanometersDelta = NanometersMax - NanometersMin
        ' ---------------------------------------------------------------------
        NmStart = NanometersMin + NanometersDelta * SrcX0 / SrcW  '光谱的起始波长，跟图像采集区域挂钩了
        NmEnd = NanometersMax - NanometersDelta * (1.0F - CSng(SrcX0 + SrcDX - 1) / SrcW)
        ' ---------------------------------------------------------------------
        NmStartDiv = 10 * CInt(NmStart / 10.0F)
        ' ---------------------------------------------------------------------
        SetNmCoeff()
    End Sub

    Private Sub Spectrometer_SetDestParams() '设置目标参数 DestPbox As PictureBox = Form_Main.PBox_Spectrum, dest - destination
        DestW = DestPbox.Image.Width         'DestPbox As PictureBox = Form_Main.PBox_Spectrum
        DestH = DestPbox.Image.Height
        DestRight = DestW - 1
        DestBottom = DestH - 2
        SetNmCoeff()
    End Sub

    Private Sub SetNmCoeff()
        If NmEnd > NmStart Then
            NmCoeff = DestW / (NmEnd - NmStart) '光谱的缩放系数，将DestW与波长挂钩
        Else
            NmCoeff = 1
        End If
    End Sub

    ' ================================================================================
    '  REFERENCE - SET RESET
    ' ================================================================================
    Private Reference(-1) As Single
    Friend Sub Spectrometer_SetReference()
        Reference = CType(SpecArrayFiltered.Clone, Single())
        For i As Int32 = 0 To SrcDX - 1
            If Reference(i) < 1 Then Reference(i) = 99999
        Next
        ReferenceScale = True
    End Sub
    Friend Sub Spectrometer_ResetReference()
        Form_Main.btn_Reference.Checked = False
        ReDim Reference(-1)
        ReferenceScale = False
    End Sub

    ' ================================================================================
    '  PROCESS CAPTURED IMAGE
    ' ================================================================================
    Friend Sub ProcessCapturedImage()
        ' -----------------------------------------------------------------------
        If Not Capture_NewImageIsReady Then Return
        Dim t As PrecisionTimer = New PrecisionTimer
        ' ----------------------------------------------------------------------- LOAD CAPTURED IMAGE
        SrcImage = Capture_Image
        If SrcImage Is Nothing Then Return
        ' ----------------------------------------------------------------------- CAPTURE IMAGE NOT READY
        Capture_NewImageIsReady = False
        ' ----------------------------------------------------------------------- FLIP
        If Flip Then
            SrcImage.RotateFlip(RotateFlipType.RotateNoneFlipX)
        Else
            'SrcImage.RotateFlip(RotateFlipType.RotateNoneFlipNone)
            'SrcImage.RotateFlip(RotateFlipType.RotateNoneFlipY)
        End If
        ' ----------------------------------------------------------------------- PREPARE BITMAP
        SrcBmp = CType(SrcImage, Bitmap)
        ' ----------------------------------------------------------------------- SET PARAMS
        If SrcImage.Width <> SrcW Or SrcImage.Height <> SrcH Then
            Spectrometer_SetSourceParams()
        End If

        '这是最最关键的函数，EXTRACT SPECTRUM
        BitmapToSpectrum()

        'SHOW AREA，绘制橘色选择框
        Dim SrcGraphics As Graphics = Graphics.FromImage(SrcBmp)
        SrcGraphics.DrawRectangle(SrcPen, SrcX0, SrcY0, SrcDX, SrcDY) 'SrcPen = New Pen(Color.FromArgb(200, 120, 0), SrcH \ 100)

        ' SHOW SOURCE IMAGE
        Form_Main.PBox_Camera.Image = SrcImage

        'IMAGE INFO，图像的尺寸和帧率
        Form_Main.Label_Resolution.Text = Capture_Image.Width.ToString & " x " & Capture_Image.Height.ToString
        Form_Main.Label_FramesPerSec.Text = Capture_FramesPerSecond.ToString("0") & " fps"
        Form_Main.Label_Millisec.Text = (t.GetTimeMicrosec / 1000.0F).ToString("0") & " mS"
        ' -----------------------------------------------------------------------
        ShowSpectrumGraph()
        ' ----------------------------------------------------------------------- 
        MeterBar1.SetValue(MaxValue / 256.0F)
        ' 左下角的状态栏， STATUS BAR INDICATIONS
        If Not CursorInside Then
            If MaxValue > 0.11 Then
                Dim nm As Int32 = CInt(X_To_Nanometers((MaxValueX * DestW) \ SrcDX)) 'NmStart + (x / NmCoeff)，NmCoeff = DestW / (NmEnd - NmStart)
                If nm <> MaxNanometers_OldValue Then
                    MaxNanometers_OldValue = nm
                    Form_Main.Label_MaxPeak.Text = "Max: " & nm.ToString & " nm"
                End If
            Else
                MaxNanometers_OldValue = -1
                Form_Main.Label_MaxPeak.Text = ""
            End If
        End If

    End Sub

    Private Sub BitmapToSpectrum()
        If SrcBmp Is Nothing Then Return
        Dim SourceData As BitmapData = SrcBmp.LockBits(New Rectangle(SrcX0, SrcY0, SrcDX, SrcDY), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb)
        Dim SourceStride As Int32 = SourceData.Stride  '这里应该是  SrcDX * 3
        Dim byteCount As Integer = (SourceData.Stride * SourceData.Height) '就是所有的像素数量
        Dim bmpBytes(byteCount - 1) As Byte
        Try
            Marshal.Copy(SourceData.Scan0, bmpBytes, 0, byteCount)  '复制到这个数组bmpBytes[]中来，组成了一个三维矩阵
        Catch
        End Try
        SrcBmp.UnlockBits(SourceData)

        Dim sumr As Int32 = 0
        Dim sumg As Int32 = 0
        Dim sumb As Int32 = 0
        Dim disp As Integer 'display
        Dim v As Single
        For x As Int32 = 0 To SrcDX - 1  '针对每个像素，来处理Y的值
            sumr = 0                    '每个pixel中每个颜色通道的取值范围是0 - 255，也就是2^8=256
            sumg = 0                    '
            sumb = 0
            disp = x * 3  '每一个像素是占用了3个字节，所以需要每隔3个字节取一次数据
            For y As Int32 = 0 To SrcDY - 1
                sumr += bmpBytes(disp + 2)
                sumg += bmpBytes(disp + 1)
                sumb += bmpBytes(disp)      'opencv 里图像的存储为 BGR 格式，刚好和现在流行的 RGB 反过来了
                disp = disp + SourceStride  '这个地方控制换行了，换到下一行了
            Next

            '相当于对每一列像素的三个颜色通道的强度进行求和，Kred = 1.0F / SrcDY，这个系数已经除以SrcDY了
            v = sumr * Kred + sumg * Kgreen + sumb * Kblue      '所以v的取值范围就是0-765，255x3 = 765

            SpecV(x) = v

            If v > SpecArray(x) Then   '需要处理的是这个数组SpecArray(x)里面的数据
                SpecArray(x) += (v - SpecArray(x)) * KSpeedUP 'txt_RisingSpeed KSpeedUP = SpeedUP / 100.0F，相当于取pixel前后两次数值的平均值
                '---------------------------------------------这个数组的范围还是按照最大值765计算
            Else
                SpecArray(x) += (v - SpecArray(x)) * KSpeedDW
            End If
        Next
        AddFilter()
        AddReference()
    End Sub

    Private Sub AddFilter() 'Noise filtering of the graph.
        MaxValue = 0.1
        Dim v, vnew As Single 'v - Value
        For i As Int32 = 0 To SrcDX - 1
            vnew = SpecArray(i)
            ' ----------------------------------------- filter，KFilter = (100 - Filter) / 100.0F + 0.1F
            v += (vnew - v) * KFilter
            ' ----------------------------------------- store filtered value
            SpecArrayFiltered(i) = v    '这个数组的范围还是按照最大值765计算，可以适当扩大到800？
        Next
        For i As Int32 = SrcDX - 1 To 0 Step -1
            vnew = SpecArray(i)
            ' ----------------------------------------- filter
            v += (vnew - v) * KFilter
            ' ----------------------------------------- add up and down filter passes
            'SpecArrayFiltered(i) += v
            SpecArrayFiltered(i) = v
            ' ----------------------------------------- update Max
            If SpecArrayFiltered(i) > MaxValue Then '在这个地方获取峰值以及峰值波长，显示在软件的左下角
                MaxValue = SpecArrayFiltered(i)
                MaxValueX = i
            End If
        Next
    End Sub

    Private Sub AddReference() '设置参考光谱，就是测吸收光谱的时候用的
        If Reference.Length = 0 Then Return
        Dim v As Single
        For i As Int32 = 0 To SrcDX - 1
            v = SpecArrayFiltered(i)
            ' ----------------------------------------- reference
            v = v * 0.9F * MaxValue / Reference(i)
            If v > MaxValue Then v = MaxValue
            ' ----------------------------------------- store value
            SpecArrayFiltered(i) = v
        Next
    End Sub


    Friend Sub ShowSpectrumGraph()
        ' --------------------------------------------------------------------
        InitPictureboxImage(DestPbox)
        If DestPbox.Image Is Nothing Then Return
        If DestPbox.Image.Width <> DestW Or DestPbox.Image.Height <> DestH Then
            Spectrometer_SetDestParams()
            gfx = Graphics.FromImage(DestPbox.Image)   'Creates a new Graphics from the specified Image.
        End If
        ' --------------------------------------------------------------------
        gfx.Clear(Color.AliceBlue) 'Clears the entire drawing surface and fills it with the specified background color.
        Dim x As Single
        Dim y As Single
        Dim drawText As Boolean
        ' --------------------------------------------------------------------- scale X，X轴比例尺，10个一格
        For i As Int32 = NmStartDiv To CInt(NmEnd) Step 10    'NmStartDiv = 10 * CInt(NmStart / 10.0F)，转换成整数，还是绘制光谱的起始波长
            drawText = False
            x = (i - NmStart) * NmCoeff
            If i Mod 100 = 0 Then
                gfx.DrawLine(ScalePen1, x, 15, x, DestBottom)   'ScalePen1 = New Pen(Color.FromArgb(140, 140, 140))
                'Draws a line connecting the two points specified by the coordinate pairs.
                drawText = True
            ElseIf i Mod 50 = 0 Then
                gfx.DrawLine(ScalePen2, x, 15, x, DestBottom)   'ScalePen2 = New Pen(Color.FromArgb(200, 200, 200))
                If NmCoeff > 2 Then drawText = True
            Else
                gfx.DrawLine(ScalePen3, x, 15, x, DestBottom)   'ScalePen3 = New Pen(Color.FromArgb(220, 220, 220))
                If NmCoeff > 5 Then drawText = True
            End If
            If drawText Then
                gfx.DrawString(i.ToString, ScaleFont, Brushes.Black, x - 4, 1)  'ScaleFont = New Font("Arial", 8)
            End If
        Next
        ' --------------------------------------------------------------------- scale Y
        If ReferenceScale Then
            For i As Int32 = 0 To 110 Step 5
                y = DestBottom - i / 110.0F * (DestBottom - 15)
                If i = 100 Then
                    gfx.DrawLine(Pens.Orange, 0, y, DestRight, y)
                ElseIf i = 50 Then
                    gfx.DrawLine(Pens.Orange, 0, y, DestRight, y)
                ElseIf i Mod 10 = 0 Then
                    gfx.DrawLine(ScalePen2, 0, y, DestRight, y)
                Else
                    gfx.DrawLine(ScalePen3, 0, y, DestRight, y)
                End If
            Next
        Else
            For i As Int32 = 0 To 100 Step 5
                y = DestBottom - i / 100.0F * (DestBottom - 15)
                If i = 100 Then
                    gfx.DrawLine(Pens.YellowGreen, 0, y, DestRight, y)
                ElseIf i = 50 Then
                    gfx.DrawLine(Pens.YellowGreen, 0, y, DestRight, y)
                ElseIf i Mod 10 = 0 Then
                    gfx.DrawLine(ScalePen2, 0, y, DestRight, y)
                Else
                    gfx.DrawLine(ScalePen3, 0, y, DestRight, y)
                End If
            Next
        End If
        '上面的部分是画网格线

        ' --------------------------------------------------------------------- graph vars
        Dim oldx As Single = -99
        Dim oldy As Single = 0
        '每一个像素分配的X轴长度，DestW = DestPbox.Image.Width, DestRight = DestW - 1, DestBottom = DestH - 2
        kx = CSng(DestRight) / (SpecArrayFiltered.Length - 1)

        '每一个相对强度分配的Y轴高度,修改的话，把最大值设为255，光照强度的范围就是从0-255，当然，需要减去背景噪音
        ky = (DestBottom - 15) / MaxV   '前值是MaxValue
        ' --------------------------------------------------------------------- graph color fill
        Dim xnew As Int32
        Dim xold As Int32
        Dim x3 As Int32
        Dim y1 As Single
        Dim y2 As Single
        If UseColors Then
            For i As Int32 = 0 To SpecArrayFiltered.Length - 1
                xnew = CInt(BinToX(i))                                              'i * kx，实际的波长
                If xnew = xold + 1 Then                                              '1nm对应1个像素点
                    y = SpecArrayFiltered(CInt(i)) * ky                             'y就是绘图时纵坐标的值，代表强度，从0-（DestBottom - 15）
                    If y > 2 Then                                                    '就是有一点点强度就算，就开始绘制
                        Pen_Graph.Color = WavelengthToColor(X_To_Nanometers(xnew))  'X_To_Nanometers = NmStart + (xnew / NmCoeff), NmCoeff = DestW / (NmEnd - NmStart)
                        gfx.DrawLine(Pen_Graph, xnew, DestBottom, xnew, DestBottom - y)
                    End If
                ElseIf xnew > xold Then
                    Pen_Graph.Color = WavelengthToColor(X_To_Nanometers(xnew))
                    y1 = SpecArrayFiltered(CInt(i - 1)) * ky
                    y2 = SpecArrayFiltered(CInt(i)) * ky
                    If y1 > 2 Or y2 > 2 Then
                        For x3 = xold + 1 To xnew                               '这个地方是为了每nm都有一个光照强度值，可能是为了更连续吧
                            y = y1 + (y2 - y1) * (x3 - xold) / (xnew - xold)
                            gfx.DrawLine(Pen_Graph, x3, DestBottom, x3, DestBottom - y)
                        Next
                    End If
                End If
                xold = xnew
            Next
        End If
        ' --------------------------------------------------------------------- graph black line
        Pen_Graph.Color = Color.FromArgb(0, 70, 0)
        For i As Int32 = 0 To SpecArrayFiltered.Length - 1
            x = BinToX(i)
            y = DestBottom - SpecArrayFiltered(i) * ky
            gfx.DrawLine(Pen_Graph, oldx, oldy, x, y)
            'gfx.DrawLine(Pen_Graph, oldx + 1, oldy, x + 1, y) ' thick line
            If NmCoeff > 10 Then
                gfx.DrawRectangle(Pen_Graph, x - 1, y - 1, 3, 3) ' points
            End If
            oldx = x
            oldy = y
        Next
        ' --------------------------------------------------------------------- 
        If TrimScale Then
            MarkTrimPoint(TrimPoint1)
            MarkTrimPoint(TrimPoint2)
        End If
        MarkAllPeaks()
        ' --------------------------------------------------------------------- 
        DestPbox.Invalidate()
    End Sub

    Friend Function X_To_Nanometers(ByVal x As Single) As Single
        If NmCoeff = 0 Then Return 0
        Return NmStart + (x / NmCoeff)
    End Function

    Friend Function X_From_Nanometers(ByVal nm As Single) As Single
        Return (nm - NmStart) * NmCoeff
    End Function

    Friend Function BinToX(ByVal bin As Int32) As Single
        Return bin * kx
    End Function

    Friend Function XtoBin(ByVal x As Single) As Int32
        If kx = 0 Then Return 0
        Return CInt(x / kx)
    End Function


    Private Font_Peaks As Font = New Font("Arial", 9)
    Private Pen_Trim1 As Pen = New Pen(Color.White)

    Private Sub MarkTrimPoint(ByVal nm As Single)   '红色外框，填充黄色
        Pen_Trim1.DashStyle = DashStyle.Dot
        Dim x As Single
        Dim w As Int32 = 26
        x = X_From_Nanometers(nm)               'Single  Return (nm - NmStart) * NmCoeff
        If x >= 0 And x < DestW Then
            Dim s As String = nm.ToString("0")
            If s.Length > 3 Then w = 34
            gfx.DrawLine(Pens.Black, x, 16, x, DestH)
            gfx.DrawLine(Pen_Trim1, x, 16, x, DestH)
            gfx.FillRectangle(Brushes.Yellow, x - 14, 0, w, 14)
            gfx.DrawRectangle(Pens.Red, x - 14, 0, w, 14)
            gfx.DrawString(s, Font_Peaks, Brushes.Black, x - 13, 0)
        End If
    End Sub

    Private Sub MarkPeak(ByVal bin As Int32, ByVal IsPeak As Boolean)  '显示峰值和峰谷
        Dim x As Single
        Dim y1 As Int32
        Dim y2 As Int32
        Dim w As Int32 = 26
        If bin > 0 Then
            x = BinToX(bin)
            If x >= 0 And x < DestW Then
                Dim s As String = X_To_Nanometers(x).ToString("0")
                If s.Length > 3 Then w = 34
                y1 = 15 + CInt((DestH - 15) * (1 - SpecArrayFiltered(bin) / MaxV)) '前值是MaxValue
                If IsPeak Then                                      '被MarkAllPeaks()调用，峰值的情况，红线
                    y2 = y1 - 20                                    '文字显示的位置
                    If y2 < DestH - 50 Then y2 = DestH - 20
                    gfx.DrawLine(Pens.Red, x, y1 + 1, x, DestH)
                Else                                                 '被MarkAllPeaks()调用，峰谷的情况，绿线
                    y2 = 30                                          '文字显示的位置
                    gfx.DrawLine(Pens.Green, x, 40, x, y1 - 3)
                End If
                gfx.FillRectangle(Brushes.Yellow, x - 14, y2, w, 14)
                gfx.DrawRectangle(Pens.Green, x - 14, y2, w, 14)
                gfx.DrawString(s, Font_Peaks, Brushes.Black, x - 13, y2) 'Private Font_Peaks As Font = New Font("Arial", 9)
            End If
        End If
    End Sub

    Private Sub MarkAllPeaks()
        Dim delta As Int32 = (20 * SrcDX) \ DestW
        If delta < 2 Then delta = 2
        Dim v As Single
        Dim valid As Boolean
        For i As Int32 = delta To SpecArrayFiltered.Length - delta - 1
            v = SpecArrayFiltered(i)
            If ShowPeaks Then
                If v > SpecArrayFiltered(i + 1) AndAlso v > SpecArrayFiltered(i - 1) AndAlso v * 100 > MaxValue Then
                    valid = True
                    For d As Int32 = 2 To delta
                        If v < SpecArrayFiltered(i + d) OrElse v < SpecArrayFiltered(i - d) Then
                            valid = False
                            Exit For
                        End If
                    Next
                    If valid Then MarkPeak(i, True)
                End If
            End If
            If ShowDips Then
                If v < SpecArrayFiltered(i + 1) AndAlso v < SpecArrayFiltered(i - 1) AndAlso v * 10000000 > MaxValue Then
                    valid = True
                    For d As Int32 = 2 To delta
                        If v > SpecArrayFiltered(i + d) OrElse v > SpecArrayFiltered(i - d) Then
                            valid = False
                            Exit For
                        End If
                    Next
                    If valid Then MarkPeak(i, False)
                End If
            End If
        Next
    End Sub

    ' ======================================================================================
    '  SPECTRUM FILE
    ' ======================================================================================
    Friend SpectrumFileSeparator As String = vbTab

    Friend Function GetSpectrumText() As String   '导出数据表
        Dim GCI As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim s As String = ""
        If SpectrumFileSeparator = vbTab Then
            s += " nm 原始灰度值 积分灰度值 滤波灰度值" + vbCrLf
            s += "原始灰度值和积分灰度值主要是调试用，以最后一列的滤波灰度值为最终结果" + vbCrLf
            For i As Int32 = 0 To SpecArrayFiltered.Length - 1
                s += X_To_Nanometers(BinToX(i)).ToString("0.0", GCI) + vbTab + CalcPercentual(SpecV(i)).ToString("0.0", GCI) +
                    vbTab + CalcPercentual(SpecArray(i)).ToString("0.0", GCI) + vbTab + CalcPercentual(SpecArrayFiltered(i)).ToString("0.0", GCI) _
                    + vbCrLf
            Next
        Else
            s += "Nanometers   Percentual" + vbCrLf
            s += "-----------------------" + vbCrLf
            For i As Int32 = 0 To SpecArrayFiltered.Length - 1
                s += X_To_Nanometers(BinToX(i)).ToString("0.0", GCI) + SpectrumFileSeparator + CalcPercentual(SpecArrayFiltered(i)).ToString("0.0", GCI).PadLeft(12) + vbCrLf
            Next
        End If
        Return s
    End Function

    Private Function CalcPercentual(ByVal value As Single) As Single
        Dim perc As Single
        If ReferenceScale Then
            perc = value * 110 / MaxValue
        Else
            'perc = value * 100 / MaxValue
            perc = value * 100 / 100
        End If
        Return perc
    End Function

End Module
