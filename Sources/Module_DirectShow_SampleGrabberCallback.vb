
Imports DirectShowLib

Friend Class SamplegrabberCallback
    Implements ISampleGrabberCB
    '
    Friend Function BufferCB(ByVal SampleTime As Double, ByVal pBuffer As System.IntPtr, ByVal BufferLen As Integer) As Integer Implements DirectShowLib.ISampleGrabberCB.BufferCB
        '
        If BufferLen = 0 Then Return Nothing
        '
        Dim w As Int32
        Dim h As Int32
        Select Case BufferLen
            Case 57600  '160 * 120 * 3
                w = 160
                h = 120
            Case 76032  '176 * 144 * 3
                w = 176
                h = 144
            Case 168960 '320 * 176 * 3
                w = 320
                h = 176
            Case 230400 '320 * 240 * 3
                w = 320
                h = 240
            Case 304128 '352 * 288 * 3
                w = 352
                h = 288
            Case 311040 '432 * 240 * 3
                w = 432
                h = 240
            Case 470016 '544 * 288 * 3
                w = 544
                h = 288
            Case 691200 '640 * 360 * 3
                w = 640
                h = 360
            Case 921600 '640 * 480 * 3
                w = 640
                h = 480
            Case 938496 '752 * 416 * 3
                w = 752
                h = 416
            Case 1036800 '720 * 480 * 3  (480p)
                w = 720
                h = 480
            Case 1075200 '800 * 448 * 3
                w = 800
                h = 448
                'Case 1244160 '720 * 576 * 3  (576p)
                '    w = 720
                '    h = 576
            Case 1244160 '864 * 480 * 3
                w = 864
                h = 480
            Case 1440000 '800 * 600 * 3
                w = 800
                h = 600
            Case 1566720 '960 * 544 * 3
                w = 960
                h = 544
            Case 1769472 '1024 * 576 * 3
                w = 1024
                h = 576
            Case 2073600 '960 * 720 * 3
                w = 960
                h = 720
            Case 2330112 '1184 * 656 * 3
                w = 1184
                h = 656
            Case 2764800 '1280 * 720 * 3  (720p)
                w = 1280
                h = 720
            Case 3207168 '1392 * 768 * 3
                w = 1392
                h = 768
            Case 3686400 '1280 * 960 * 3
                w = 1280
                h = 960
            Case 3753984 '1504 * 832 * 3
                w = 1504
                h = 832
            Case 3932160 '1280 * 1024 * 3
                w = 1280
                h = 1024
            Case 4300800 '1600 * 896 * 3
                w = 1600
                h = 896
            Case 4930560 '1712 * 960 * 3
                w = 1712
                h = 960
            Case 5419008 '1792 * 1008 * 3
                w = 1792
                h = 1008
            Case 5760000 '1600 * 1200 * 3
                w = 1600
                h = 1200
            Case 6220800 '1920 * 1080 * 3  (1080p)
                w = 1920
                h = 1080
            Case 14745600 '2560 * 1920 * 3
                w = 2560
                h = 1920
            Case 13063680 '2592 * 1680 * 3
                w = 2592
                h = 1680
            Case 23040000 '3200 * 2400 * 3
                w = 3200
                h = 2400
        End Select
        '
        Dim nBytes As Int32 = w * h * 3
        Dim stride As Int32 = w * 3
        '
        If BufferLen = nBytes And Not Capture_NewImageIsReady Then
            'If BufferLen = nBytes Then
            Capture_Image = New Bitmap(w, h, stride, Imaging.PixelFormat.Format24bppRgb, pBuffer)
            Capture_NewImageIsReady = True
        End If
        '
        Static oldSampleTime As Double
        Dim millisec As Double = SampleTime - oldSampleTime
        oldSampleTime = SampleTime
        If millisec > 0 Then
            Capture_FramesPerSecond += (1 / millisec - Capture_FramesPerSecond) / 4
        End If

        'Beep()
    End Function
    '
    Private Function SampleCB(ByVal SampleTime As Double, ByVal pSample As DirectShowLib.IMediaSample) As Integer Implements DirectShowLib.ISampleGrabberCB.SampleCB
        'Debug.Print(SampleTime.ToString & " " & pSample.GetActualDataLength.ToString)
        'Marshal.ReleaseComObject(pSample)  ' <<< without this release the process locks
        'Return 0
    End Function
    '
End Class
