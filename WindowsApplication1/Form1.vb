Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb.OleDbConnection
Imports System.Data.OleDb
Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Foldername As String 'folder namne
        Dim S878 As New System.Text.UTF8Encoding 'Creat notpad 
        Dim BP, BS, B1, B2, S1, S2, S3, S4, A1, A2, A3, A4, improtFile As String 'txt paremeter
        Dim xlApp As Microsoft.Office.Interop.Excel.Application 'Open App
        Dim xlBook As Microsoft.Office.Interop.Excel.Workbook   ' Open work book
        Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet 'Sheet
        Dim Col, Row As Integer

        
        'Start in Excel 

        improtFile = "C:\TxtGenerator\TC_Sagi ISC_OfflineModuleCombinationTest_V0.3.xls"
        xlApp = New Microsoft.Office.Interop.Excel.Application
        If Not FileIO.FileSystem.FileExists(improtFile) Then 'to check the file is enable to read
            MsgBox("檔案不存在！")
            Return
        End If

        xlBook = xlApp.Workbooks.Open(improtFile)
        xlSheet = xlBook.Worksheets(2) ' the second sheet 

        

        For Row = 3 To 45 Step 14 ' row
            For Col = 4 To 108 Step 1 'column
                '////////////////////////////////////////////////////////// to make sure the folder have creat or not
                Foldername = xlSheet.Cells(Col, Row).Value
                If Not IO.Directory.Exists("D:\TestResult\" & Foldername) Then

                    IO.Directory.CreateDirectory("D:\TestResult\" & Foldername)
                End If


                '///////////////////////////////////////////////////////////



                BP = "" '//backplane
                BS = ""
                B1 = ""
                B2 = ""
                S1 = ""
                S2 = ""
                S3 = ""
                S4 = ""
                A1 = "" 'ISC_SWITCH_RECIPE_CONFIG
                A2 = "" '//ISC_MAINFRAME_MODEL
                A3 = "" '//ISC_BP_CONNECT_STATUS
                A4 = "" '//ISC_BACKPLANE_TYPE

                BP = xlSheet.Cells(Col, Row - 1).Value 'BP of col

                If BP = "0" Then
                    A2 = """ISC_MAINFRAME_MODEL"" = ""Taurus-SSS_TS-8503"""
                    A4 = """ISC_BACKPLANE_TYPE"" = dword:00000000"

                ElseIf BP = "1" Then

                    A2 = """ISC_MAINFRAME_MODEL"" = ""Taurus-SSS_TS-8501"""
                    A4 = """ISC_BACKPLANE_TYPE"" = dword:00000001"
                ElseIf BP = "2" Then

                    A2 = """ISC_MAINFRAME_MODEL"" = ""Taurus-SSS_TS-8502"""
                    A4 = """ISC_BACKPLANE_TYPE"" = dword:00000002"

                End If


                BS = xlSheet.Cells(Col, Row + 2).Value 'BS of col
                B1 = xlSheet.Cells(Col, Row + 3).Value 'B1 of col
                B2 = xlSheet.Cells(Col, Row + 4).Value  'B2 of col


                If BS = "1" And B1 = "0" And B2 = "0" Then

                    A3 = """ISC_BP_CONNECT_STATUS"" = ""1,1,1,1,"""


                ElseIf BS = "0" And B1 = "0" And B2 = "0" Then

                    A3 = """ISC_BP_CONNECT_STATUS"" = ""0,0,0,0,"""


                ElseIf BS = "0" And B1 = "1" And B2 = "0" Then

                    A3 = """ISC_BP_CONNECT_STATUS"" = ""1,1,0,0,"""


                ElseIf BS = "0" And B1 = "0" And B2 = "1" Then

                    A3 = """ISC_BP_CONNECT_STATUS"" = ""0,0,2,2,"""


                ElseIf BS = "0" And B1 = "1" And B2 = "1" Then

                    A3 = """ISC_BP_CONNECT_STATUS"" = ""1,1,2,2,"""
                End If



                S1 = xlSheet.Cells(Col, Row + 5).Value 'S1 of col
                S2 = xlSheet.Cells(Col, Row + 6).Value 'S2 of col
                S3 = xlSheet.Cells(Col, Row + 7).Value 'S3 of col
                S4 = xlSheet.Cells(Col, Row + 8).Value 'S4 of col

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''121K

                If S1 = "121K" And S2 = "Blank" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,-1:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "121K" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,0:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "121K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,0:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "Blank" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,-1:A01.02,0:A01.02,"""

                ElseIf S1 = "121K" And S2 = "121K" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,0:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "121K" And S2 = "Blank" And S3 = "121K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,-1:A01.02,0:A01.02,-1:A01.02,"""

                ElseIf S1 = "121K" And S2 = "Blank" And S3 = "Blank" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,-1:A01.02,-1:A01.02,0:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "121K" And S3 = "121K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,0:A01.02,0:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "121K" And S3 = "Blank" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,0:A01.02,-1:A01.02,0:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "121K" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,0:A01.02,0:A01.02,"""

                ElseIf S1 = "121K" And S2 = "121K" And S3 = "121K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,0:A01.02,0:A01.02,-1:A01.02,"""

                ElseIf S1 = "121K" And S2 = "121K" And S3 = "Blank" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,0:A01.02,-1:A01.02,0:A01.02,"""

                ElseIf S1 = "121K" And S2 = "Blank" And S3 = "121K" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,-1:A01.02,0:A01.02,0:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "121K" And S3 = "121K" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,0:A01.02,0:A01.02,0:A01.02,"""

                ElseIf S1 = "121K" And S2 = "121K" And S3 = "121K" And S4 = "121K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""0:A01.02,0:A01.02,0:A01.02,0:A01.02,"""
                End If











                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''122K
                If S1 = "122K" And S2 = "Blank" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,-1:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "122K" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,14:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "122K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,14:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "Blank" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,-1:A01.02,14:A01.02,"""

                ElseIf S1 = "122K" And S2 = "122K" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,14:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "122K" And S2 = "Blank" And S3 = "122K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,-1:A01.02,14:A01.02,-1:A01.02,"""

                ElseIf S1 = "122K" And S2 = "Blank" And S3 = "Blank" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,-1:A01.02,-1:A01.02,14:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "122K" And S3 = "122K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,14:A01.02,14:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "122K" And S3 = "Blank" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,14:A01.02,-1:A01.02,14:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "122K" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,14:A01.02,14:A01.02,"""

                ElseIf S1 = "122K" And S2 = "122K" And S3 = "122K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,14:A01.02,14:A01.02,-1:A01.02,"""

                ElseIf S1 = "122K" And S2 = "122K" And S3 = "Blank" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,14:A01.02,-1:A01.02,14:A01.02,"""

                ElseIf S1 = "122K" And S2 = "Blank" And S3 = "122K" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,-1:A01.02,14:A01.02,14:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "122K" And S3 = "122K" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,14:A01.02,14:A01.02,14:A01.02,"""

                ElseIf S1 = "122K" And S2 = "122K" And S3 = "122K" And S4 = "122K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""14:A01.02,14:A01.02,14:A01.02,14:A01.02,"""

                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''123K
                If S1 = "123K" And S2 = "Blank" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,-1:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "123K" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,6:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "123K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,6:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "Blank" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,-1:A01.02,6:A01.02,"""

                ElseIf S1 = "123K" And S2 = "123K" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,6:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "123K" And S2 = "Blank" And S3 = "123K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,-1:A01.02,6:A01.02,-1:A01.02,"""

                ElseIf S1 = "123K" And S2 = "Blank" And S3 = "Blank" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,-1:A01.02,-1:A01.02,6:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "123K" And S3 = "123K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,6:A01.02,6:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "123K" And S3 = "Blank" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,6:A01.02,-1:A01.02,6:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "123K" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,6:A01.02,6:A01.02,"""

                ElseIf S1 = "123K" And S2 = "123K" And S3 = "123K" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,6:A01.02,6:A01.02,-1:A01.02,"""

                ElseIf S1 = "123K" And S2 = "123K" And S3 = "Blank" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,6:A01.02,-1:A01.02,6:A01.02,"""

                ElseIf S1 = "123K" And S2 = "Blank" And S3 = "123K" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,-1:A01.02,6:A01.02,6:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "123K" And S3 = "123K" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,6:A01.02,6:A01.02,6:A01.02,"""

                ElseIf S1 = "123K" And S2 = "123K" And S3 = "123K" And S4 = "123K" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""6:A01.02,6:A01.02,6:A01.02,6:A01.02,"""

                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''1220
                If S1 = "1220" And S2 = "Blank" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,-1:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "1220" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,7:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "1220" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,7:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "Blank" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,-1:A01.02,7:A01.02,"""

                ElseIf S1 = "1220" And S2 = "1220" And S3 = "Blank" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,7:A01.02,-1:A01.02,-1:A01.02,"""

                ElseIf S1 = "1220" And S2 = "Blank" And S3 = "1220" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,-1:A01.02,7:A01.02,-1:A01.02,"""

                ElseIf S1 = "1220" And S2 = "Blank" And S3 = "Blank" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,-1:A01.02,-1:A01.02,7:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "1220" And S3 = "1220" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,7:A01.02,7:A01.02,-1:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "1220" And S3 = "Blank" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,7:A01.02,-1:A01.02,7:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "Blank" And S3 = "1220" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,-1:A01.02,7:A01.02,7:A01.02,"""

                ElseIf S1 = "1220" And S2 = "1220" And S3 = "1220" And S4 = "Blank" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,7:A01.02,7:A01.02,-1:A01.02,"""

                ElseIf S1 = "1220" And S2 = "1220" And S3 = "Blank" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,7:A01.02,-1:A01.02,7:A01.02,"""

                ElseIf S1 = "1220" And S2 = "Blank" And S3 = "1220" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,-1:A01.02,7:A01.02,7:A01.02,"""

                ElseIf S1 = "Blank" And S2 = "1220" And S3 = "1220" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""-1:A01.02,7:A01.02,7:A01.02,7:A01.02,"""

                ElseIf S1 = "1220" And S2 = "1220" And S3 = "1220" And S4 = "1220" Then

                    A1 = """ISC_SWITCH_RECIPE_CONFIG"" = ""7:A01.02,7:A01.02,7:A01.02,7:A01.02,"""

                End If


                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''TXT info
                Using sb As System.IO.StreamWriter = New System.IO.StreamWriter("D:\TestResult\" & Foldername & " \sagi-isc.txt", False, S878)
                    ' .txt information
                    sb.WriteLine("Windows Registry Editor Version 5.00".ToString)
                    sb.WriteLine("""CompanyID""=dword:00000001".ToString)
                    sb.WriteLine("""RecipePath""=""C:\Program Files\STAr\Sagittarius\Setup XML\""".ToString)
                    sb.WriteLine("""RecipeFile""=""C:\Program Files\STAr\Sagittarius\Setup XML\""".ToString)
                    sb.WriteLine("""AnalyzerWindow""=""2""".ToString)
                    sb.WriteLine("""InstallPath""=""C:\Program Files\STAr\Sagittarius\""".ToString)
                    sb.WriteLine("""TestRunOutputPath""=""C:\STAr\Sagittarius\TestrunResults\""".ToString)
                    sb.WriteLine("""EnggResultPath""=""C:\STAr\Sagittarius\EnggUIResults\""".ToString)
                    sb.WriteLine("""GraphResultPath""=""C:\STAr\Sagittarius\GraphResults\""".ToString)
                    sb.WriteLine("""ICTResultPath""=""C:\STAr\Sagittarius\ICTResults\""".ToString)
                    sb.WriteLine("""IVUMode""=dword:00000000".ToString)
                    sb.WriteLine("""ProbeSpecFile""=""Default_ProbeSpec.xml""".ToString)
                    sb.WriteLine("""LogicalProberPins""=""48""".ToString)
                    sb.WriteLine("""GraphSpecFile""=""Default_GraphSpec.xml""".ToString)
                    sb.WriteLine("""FlowDllName""=""SagiStdFlow.dll""".ToString)
                    sb.WriteLine("""EnableDebugLogging""=dword:00000001".ToString)
                    sb.WriteLine("""Language""=""sagittariuseng.txt""".ToString)
                    sb.WriteLine("""HardwareConfig""=""localhw.xml""".ToString)
                    sb.WriteLine("""ProdIpMode""=dword:00000000".ToString)
                    sb.WriteLine("""GroupID""=dword:00000001".ToString)
                    sb.WriteLine("""RemoteDataPath""=""""".ToString)
                    sb.WriteLine("""Operation_Mode""=dword:00000000".ToString)
                    sb.WriteLine("""Boundary_Criteria""=dword:00000001".ToString)
                    sb.WriteLine("""Die_Unit""=dword:00000001".ToString)
                    sb.WriteLine("""Die_Value""=""1""".ToString)
                    sb.WriteLine("""Wafer_Unit""=dword:00000002".ToString)
                    sb.WriteLine("""Wafer_Value""=""1""".ToString)
                    sb.WriteLine("""Lot_Unit""=dword:00000003".ToString)
                    sb.WriteLine("""Lot_Value""=""1""".ToString)
                    sb.WriteLine("""Auto_Generate_Test_Prefix""=dword:00000001".ToString)
                    sb.WriteLine("""Check_Duplicate_Test_Item""=dword:00000000".ToString)
                    sb.WriteLine("""LicenceKey""=""""".ToString)
                    sb.WriteLine("""VersionInfo""=""""".ToString)
                    sb.WriteLine("""MDR_Options""=""0""".ToString)
                    sb.WriteLine("""MDR_NoOfSite""=""16""".ToString)
                    sb.WriteLine("""MDR_ProbePinsPerSite""=""24""".ToString)
                    sb.WriteLine("""MDR_MaxChuckTemp""=""200""".ToString)
                    sb.WriteLine("""MDR_HCI_MaxDUTsPerSite""=""6""".ToString)
                    sb.WriteLine("""MDR_HCI_MaxDiesPerGroup""=""16""".ToString)
                    sb.WriteLine("""MDR_HCI_MaxStressGroup_BBEnabled""=""4""".ToString)
                    sb.WriteLine("""MDR_HCI_MaxStressGroup_BBDisabled""=""6""".ToString)
                    sb.WriteLine("""MDR_NBTI_MaxDUTsPerSite""=""6""".ToString)
                    sb.WriteLine("""MDR_NBTI_MaxDiesPerGroup""=""16""".ToString)
                    sb.WriteLine("""MDR_NBTI_MaxStressGroup_BBEnabled""=""4""".ToString)
                    sb.WriteLine("""MDR_NBTI_MaxStressGroup_BBDisabled""=""6""".ToString)
                    sb.WriteLine("""MDR_TDDB_MaxDUTsPerSite""=""12""".ToString)
                    sb.WriteLine("""MDR_TDDB_MaxDiesPerGroup""=""16""".ToString)
                    sb.WriteLine("""MDR_TDDB_MaxStressGroup""=""12""".ToString)
                    sb.WriteLine("""MDR_AGOI_MaxDUTsPerSite""=""6""".ToString)
                    sb.WriteLine("""MDR_AGOI_MaxDiesPerGroup""=""16""".ToString)
                    sb.WriteLine("""MDR_AGOI_MaxStressGroup_SBEnabled""=""4""".ToString)
                    sb.WriteLine("""MDR_AGOI_MaxStressGroup_SBDisabled""=""6""".ToString)
                    sb.WriteLine("""MDR_DebugState""=""0""".ToString)
                    sb.WriteLine("""MDR_PadNameOption""=""0""".ToString)
                    sb.WriteLine("""MDR_HCI_PadNameOption""=""0""".ToString)
                    sb.WriteLine("""MDR_NBTI_PadNameOption""=""0""".ToString)
                    sb.WriteLine("""MDR_AGOI_PadNameOption""=""0""".ToString)
                    sb.WriteLine("""MDR_TDDB_PadNameOption""=""1""".ToString)
                    sb.WriteLine("""GISGraphUpdateSettings""=dword:00000000".ToString)
                    sb.WriteLine("""GISTimeDuration""=""600""".ToString)
                    sb.WriteLine("""FlashResultPath""=""C:\STAr\Sagittarius\FlashResults\""".ToString)
                    sb.WriteLine("""MTMResultPath""=""C:\STAr\Sagittarius\MTMResults\""".ToString)
                    sb.WriteLine("""MTMFileStructureOption""=dword:00000000".ToString)
                    sb.WriteLine("""ProdOpMode""=dword:00000000".ToString)
                    sb.WriteLine("""Auto_Mail_Notification""=dword:00000000".ToString)
                    sb.WriteLine("""Engg_Mail_Notification""=dword:00000000".ToString)
                    sb.WriteLine("""Flash_Mail_Notification""=dword:00000000".ToString)
                    sb.WriteLine("""ICT_Mail_Notification""=dword:00000000".ToString)
                    sb.WriteLine("""Testmail_Template""=""Default(Sample)""".ToString)
                    sb.WriteLine("""GVEFile""=""HCEM.gve""".ToString)
                    sb.WriteLine("""ICTTestDataOption""=dword:00000000".ToString)
                    sb.WriteLine("""ICTTestSetupOption""=dword:00000000".ToString)
                    sb.WriteLine("""PowerLineFreq""=dword:00000000".ToString)
                    sb.WriteLine("""LanguageType""=dword:00000000".ToString)
                    sb.WriteLine("""GroundProbePins_Mode""=dword:00000000".ToString)
                    sb.WriteLine("""GroundProbePins_TimeSpan""="" - 1.000000""".ToString)
                    sb.WriteLine("""DataExport_Enabled""=dword:00000000".ToString)
                    sb.WriteLine("""DataExport_ExeFile""=""""".ToString)
                    sb.WriteLine("""DataExport_OutputFolder""=""""".ToString)
                    sb.WriteLine("""Server_File_Mode""=dword:00000000".ToString)
                    sb.WriteLine("""Server_File_Path""=""""")
                    sb.WriteLine("""MasterIPAddress""=""""")
                    sb.WriteLine("""MasterDataResultPath""=""C:\STAr\Sagittarius\Data\""")
                    sb.WriteLine("""CIMAutomationCheck""=dword:00000000")
                    sb.WriteLine("""CIMHostIPAddress""=""0.0.0.0""")
                    sb.WriteLine("""CIMHostPort""=dword:00000000")
                    sb.WriteLine("""CIMHostConnectTimeout""=dword:00000100")
                    sb.WriteLine("""Max_Station""=dword:00000001")
                    sb.WriteLine("""Default ISS Location""=""""")
                    sb.WriteLine("""Aux Site""=""""")
                    sb.WriteLine("""Aux Zone""=""""")
                    sb.WriteLine("""Tester Instrument""=""""")
                    sb.WriteLine("""Port Mapping""=""""")
                    sb.WriteLine("""Automatically generate *.SnP file, include parameter: ""=dword:00000000")
                    sb.WriteLine("""H""=dword:00000000")
                    sb.WriteLine("""Y""=dword:00000000")
                    sb.WriteLine("""Z""=dword:00000000")
                    sb.WriteLine("""Station_1""=""0, -, -, 0.0.0.0, -, 0""")
                    sb.WriteLine("""IDEType""=dword:00000001")
                    sb.WriteLine("""SystemSource_Enabled""=dword:00000000")
                    sb.WriteLine("""DataExport_SystemSourceFolder""=""""")
                    sb.WriteLine("""Algo_DLL_Mode""=dword:00000000")
                    sb.WriteLine("""Prober_WaferMap_File_Mode""=dword:00000000")
                    sb.WriteLine("""Prober_WaferMap_File_Path""=""""")
                    sb.WriteLine("""Activate_Signal_Tower""=dword:00000000")
                    sb.WriteLine("""Light_Location_Top""=""Red""")
                    sb.WriteLine("""Light_Location_Middle""=""Amber""")
                    sb.WriteLine("""Light_Location_Bottom""=""Green""")
                    sb.WriteLine("""System_Power_Off""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""System_Power_On""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Emo_On""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Sagittarius_StartUp""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Sagittarius_Shutdown_Ok""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Sagittarius_Shutdown_Ng""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Test_Run_Ok""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Test_Run_Ng""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Test_Complete""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Test_Abort""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Error_Clear""=""0, 0, 0, 0, 0, 0, """)
                    sb.WriteLine("""Integrate_SCI""=dword:00000000")
                    sb.WriteLine("""SCI_Type""=""""")
                    sb.WriteLine("""SCI_System_Power_Off""="" -, -, -, """)
                    sb.WriteLine("""SCI_System_Power_ON""="" -, -, -, """)
                    sb.WriteLine("""SCI_EMO_ON""="" -, -, -, """)
                    sb.WriteLine("""SCI_Sagittarius_Startup""=""1, 0, 0, """)
                    sb.WriteLine("""SCI_Sagittarius_Shutdown_OK""=""0, 0, 0, """)
                    sb.WriteLine("""SCI_Sagittarius_Shutdown_NG""="" -, -, -, """)
                    sb.WriteLine("""SCI_Test_Running_OK""="" -, 0, -, """)
                    sb.WriteLine("""SCI_Test_Running_NG""="" -, 1, -, """)
                    sb.WriteLine("""SCI_Test_Complete""="" -, -, -, """)
                    sb.WriteLine("""SCI_Test_Abort""="" -, -, -, """)
                    sb.WriteLine("""SCI_Error_Clear""="" -, 0, -, """)
                    sb.WriteLine("""SCI_High_V""="" -, -, 1, """)
                    sb.WriteLine("""SCI_N_High_V""="" -, -, 0, """)
                    sb.WriteLine("""ISC_UPDATE_FLAG""=dword:00000000")
                    sb.WriteLine("""ISC_LAUNCH_STATUS""=dword:00000000")
                    sb.WriteLine("""ISC_STARTUP_FLAG""=dword:00000000")
                    sb.WriteLine("""ISC_COMM_LOGGING_FLAG""=dword:00000001")
                    sb.WriteLine("""ISC_KS_B2200_CHECKED_FLAG""=dword:00000001")
                    sb.WriteLine("""ISC_KS_B700_CHECKED_FLAG""=dword:00000000")
                    sb.WriteLine("""ISC_SLAVE_ENABLED_FLAG""=dword:00000001")
                    sb.WriteLine("""ISC_SLAVE_INTERFACE""=""GPIB""")
                    sb.WriteLine("""ISC_GPIB_ADDRESS""=""22""")
                    sb.WriteLine("""ISC_LOG_COUNT""=""10""")
                    sb.WriteLine("""ISC_VERSION""=""1.0.0""")
                    sb.WriteLine(A1)
                    sb.WriteLine("""ISC_LOG_TIME_INTERVAL""=""10""")
                    sb.WriteLine(A2)
                    sb.WriteLine(A3)
                    sb.WriteLine(A4)
                    sb.WriteLine("""ISC_CONNECT_RULE_STATUS""=""0, 0, 0, 0, """)
                    sb.WriteLine("""ISC_LOG_REALTIME""=dword:00000001")
                    sb.WriteLine("""FIRMWARE_VER""=""A00.00""")
                    sb.WriteLine("""ISC_DEVICE_ID_MAINFRAME""=""PXI1::0::INSTR""")
                    sb.WriteLine("""ISC_DEVICE_ID_SLOT_1""=""PXI1::0::INSTR""")
                    sb.WriteLine("""ISC_DEVICE_ID_SLOT_2""=""PXI1::0::INSTR""")
                    sb.WriteLine("""ISC_DEVICE_ID_SLOT_3""=""PXI1::0::INSTR""")
                    sb.WriteLine("""ISC_DEVICE_ID_SLOT_4""=""PXI1::0::INSTR""")


                End Using
            Next
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class
