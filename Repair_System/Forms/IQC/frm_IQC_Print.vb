'############################################################################################################
'############################################################################################################

'##### 작성일자 : 2023-07-13
'##### 수정일자 : 2023-07-13
'##### 수정자   : 박시현
'##### 참고     : 
'##### 설명     : 입고된 Lot의 정보를 확인하고 YJ No.를 매칭시켜 라벨을 발행하는 폼

'############################################################################################################
'############################################################################################################

Imports C1.Win.C1FlexGrid
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Text
Imports System.Threading

Public Class frm_IQC_Print

    Dim form_Msgbox_String As String = "수입검사 및 라벨발행"
    Private WithEvents ComPort As New Ports.SerialPort
    Dim c_product_code As String
    Dim c_qty As Integer
    Dim c_fab_line As String
    Dim c_ww As String
    Dim c_option As String
    Dim change_pmic_top_marking, change_rcd_top_marking As String
    Dim before_pmic_top_marking, before_rcd_top_marking As String
    Dim part_top_marking As String
    Dim first_run As Boolean = True

    Private Sub IN_N_PRINT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        SplitContainer1.Panel1Collapsed = True

        GRID_SETTING()
        Setting_LOAD()

        Label17.Text = Format(Now, "yyyy년 MM월 dd일 입고이력")
        TB_Product_Code.Focus()
        Inspect_Result_Read()

    End Sub

    Private Sub GRID_SETTING()

        With grid_Inspect_List
            .AllowEditing = False
            .AllowFiltering = False
            .AllowSorting = AllowSortingEnum.None
            .AllowFreezing = AllowFreezingEnum.None
            .Rows(0).Height = 40
            .Rows.DefaultSize = 20
            .Cols.Count = 15
            .Cols.Fixed = 1
            .Rows.Count = 1
            .Rows.Fixed = 1
            grid_Inspect_List(0, 0) = "No"
            grid_Inspect_List(0, 1) = "검사(발행) 시간"
            grid_Inspect_List(0, 2) = "구분"
            grid_Inspect_List(0, 3) = "Product Code" & vbCrLf & "/ Part No."
            grid_Inspect_List(0, 4) = "Lot No."
            grid_Inspect_List(0, 5) = "Qty"
            grid_Inspect_List(0, 6) = "Fab Line"
            grid_Inspect_List(0, 7) = "W/W"
            grid_Inspect_List(0, 8) = "Option"
            grid_Inspect_List(0, 9) = "YJ No."
            grid_Inspect_List(0, 10) = "정량검사"
            grid_Inspect_List(0, 11) = "외관검사"
            grid_Inspect_List(0, 12) = "정보비교"
            grid_Inspect_List(0, 13) = "검사자"
            grid_Inspect_List(0, 14) = "비고"
            .AutoClipboard = True
            .Styles.Fixed.TextAlign = TextAlignEnum.CenterCenter
            .Styles.Normal.TextAlign = TextAlignEnum.CenterCenter
            '.Cols(.Cols.Count - 1).StyleNew.TextAlign = TextAlignEnum.LeftCenter
            .AutoSizeCols()
            .ShowCursor = True
            .ExtendLastCol = True
        End With

    End Sub

    Private Sub BTN_SavePrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_SavePrint.Click

        If CB_Parts_Section.Text = "Module" Then
            module_Save()
        Else
            material_Save()
        End If

    End Sub

    Private Sub material_Save()

        If Lot_Exist_Check() = True Then
            MsgBox("수입검사 완료된 Lot No.입니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        'If Not c_product_code = TB_Product_Code.Text Then
        '    MsgBox("입력된 Part No.가 다릅니다." & vbCrLf & "확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
        '    TB_Product_Code.SelectAll()
        '    TB_Product_Code.Focus()
        '    Exit Sub
        'End If

        'If Not c_qty = TB_Qty.Text Then
        '    MsgBox("입력된 수량이 다릅니다." & vbCrLf & "확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
        '    TB_Qty.SelectAll()
        '    TB_Qty.Focus()
        '    Exit Sub
        'End If

        Dim qty_result As String = "불합격"
        If RB_Qty_OK.Checked Then qty_result = "합격"
        Dim insp_result As String = "불합격"
        If RB_Insp_OK.Checked Then insp_result = "합격"
        Dim info_result As String = "불합격"
        If RB_Info_OK.Checked Then info_result = "합격"

        If CB_Print_Use.Checked = False Then
            If MsgBox("라벨 발행이 선택되지 않았습니다." & vbCrLf &
                      "계속 진행 하시겠습니까?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub
        End If

        If MsgBox("입고 검사를 완료 하시겠습니까?" & vbCrLf &
                  "정량확인 결과 : " & qty_result & vbCrLf &
                  "외관확인 결과 : " & insp_result & vbCrLf &
                  "정보비교 결과 : " & info_result, MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        '##### 라벨발행 #####
        If CB_Print_Use.Checked Then
            Dim print_check As Boolean = Label_Print_Run(CB_Parts_Section.Text)

            If print_check = False Then
                thread_LoadingFormEnd
                Thread.Sleep(100)
                MsgBox("라벨 발행 실패", MsgBoxStyle.Critical, form_Msgbox_String)
                Exit Sub
            End If
        End If

        '##### DB 저장 #####
        DB_Write("Material")

        '##### Control 초기화 #####
        TB_Product_Code.Text = String.Empty
        TB_Lot_No.Text = String.Empty
        TB_Qty.Text = String.Empty
        TB_Fab_line.Text = String.Empty
        TB_WW.Text = String.Empty
        TB_Option.Text = String.Empty
        TB_YJ_No.Text = String.Empty
        RB_Info_OK.Checked = True
        RB_Insp_OK.Checked = True
        RB_Qty_OK.Checked = True
        TB_Insp_Memo.Text = String.Empty
        tb_Before_Product.Text = String.Empty

        thread_LoadingFormEnd
        Thread.Sleep(100)
        MsgBox("저장완료.", MsgBoxStyle.Information, form_Msgbox_String)

        TB_Product_Code.SelectAll()
        TB_Product_Code.Focus()

        Inspect_Result_Read()

    End Sub

    Private Sub module_Save()

        If CB_Parts_Section.Text = "Module" And TB_YJ_No.Text = String.Empty Then
            MsgBox("YJ No.가 생성되지 않았습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If Lot_Exist_Check() = True Then
            MsgBox("수입검사 완료된 Lot No.입니다.", MsgBoxStyle.Information, form_Msgbox_String)
            Exit Sub
        End If

        If Not c_product_code = TB_Product_Code.Text Then
            MsgBox("입력된 Product Code가 다릅니다." & vbCrLf & "확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            TB_Product_Code.SelectAll()
            TB_Product_Code.Focus()
            Exit Sub
        End If

        If Not c_qty = TB_Qty.Text Then
            MsgBox("입력된 수량이 다릅니다." & vbCrLf & "확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            TB_Qty.SelectAll()
            TB_Qty.Focus()
            Exit Sub
        End If

        If Not c_fab_line = TB_Fab_line.Text And Not c_fab_line = String.Empty Then
            MsgBox("입력된 Fab Line이 다릅니다." & vbCrLf & "확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
            TB_Fab_line.SelectAll()
            TB_Fab_line.Focus()
            Exit Sub
        End If

        Dim qty_result As String = "불합격"
        If RB_Qty_OK.Checked Then qty_result = "합격"
        Dim insp_result As String = "불합격"
        If RB_Insp_OK.Checked Then insp_result = "합격"
        Dim info_result As String = "불합격"
        If RB_Info_OK.Checked Then info_result = "합격"

        If CB_Print_Use.Checked = False Then
            If MsgBox("라벨 발행이 선택되지 않았습니다." & vbCrLf &
                      "계속 진행 하시겠습니까?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub
        End If

        If MsgBox("입고 검사를 완료 하시겠습니까?" & vbCrLf &
                  "정량확인 결과 : " & qty_result & vbCrLf &
                  "외관확인 결과 : " & insp_result & vbCrLf &
                  "정보비교 결과 : " & info_result, MsgBoxStyle.Question + MsgBoxStyle.YesNo, form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        thread_LoadingFormStart("Saving...")

        '##### 라벨발행 #####
        If CB_Print_Use.Checked Then
            Dim print_check As Boolean = Label_Print_Run(CB_Parts_Section.Text)

            If print_check = False Then
                thread_LoadingFormEnd
                Thread.Sleep(100)
                MsgBox("라벨 발행 실패", MsgBoxStyle.Critical, form_Msgbox_String)
                Exit Sub
            End If
        End If

        '##### DB 저장 #####
        DB_Write("Module")

        '##### Control 초기화 #####
        TB_Product_Code.Text = String.Empty
        TB_Lot_No.Text = String.Empty
        TB_Qty.Text = String.Empty
        TB_Fab_line.Text = String.Empty
        TB_WW.Text = String.Empty
        TB_Option.Text = String.Empty
        TB_YJ_No.Text = String.Empty
        RB_Info_OK.Checked = True
        RB_Insp_OK.Checked = True
        RB_Qty_OK.Checked = True
        TB_Insp_Memo.Text = String.Empty
        tb_Before_Product.Text = String.Empty

        thread_LoadingFormEnd
        Thread.Sleep(100)
        MsgBox("저장 완료.", MsgBoxStyle.Information, form_Msgbox_String)

        TB_Product_Code.SelectAll()
        TB_Product_Code.Focus()

        Inspect_Result_Read()

    End Sub

    Private Function Label_Print_Run(ByVal label_Select As String) As Boolean

        Label_Print_Run = False

        Dim swFile As StreamWriter =
            New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

        If label_Select = "Module" Then
            Dim lotNo As String = String.Empty
            Dim rcd_information As String = tb_After_RCD.Text & " ( " & change_rcd_top_marking & " )"
            Dim pmic_information As String = tb_After_PMIC.Text & " ( " & change_pmic_top_marking & " )"
            Dim before_pmic_information As String = tb_Before_PMIC.Text & " ( " & before_pmic_top_marking & " )"
            Dim before_rcd_information As String = tb_Before_RCD.Text & " ( " & before_rcd_top_marking & " )"

            'If tb_Before_PMIC.Text = tb_After_PMIC.Text Then
            '    pmic_information = String.Empty 'PMIC 수리하지 않으므로 삭제
            'End If

            'If tb_Before_RCD.Text = tb_After_RCD.Text Then
            '    rcd_information = String.Empty 'RCD 수리하지 않으므로 삭제
            'End If

            If TB_Option.Text Like "*O" Then '숫자영 영어오
                rcd_information = String.Empty
                before_rcd_information = String.Empty
            ElseIf TB_Option.Text Like "*Q" Then

            Else
                MsgBox("수리 정보가 확실하지 않은 Option 입니다. 확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
                Return Label_Print_Run
                Exit Function
            End If

            '######################## 수입검사 박스 라벨 발행
            lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-" & TB_IQC_Char.Text & "01" & " / " & tb_After_CC.Text
            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & pmic_information & "^FS")
            swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & rcd_information & "^FS")
            swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & TB_Product_Code.Text & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^PQ" & CInt(TB_IQC_QTY.Text) & "^XZ") 'PQ : 발행수량, "ZX : 종료
            '######################## Baking JIG 라벨 발행
            lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-" & TB_Baking_Char.Text & "01" & " / " & tb_After_CC.Text
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & pmic_information & "^FS")
            swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & rcd_information & "^FS")
            swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & TB_Product_Code.Text & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^PQ" & CInt(TB_Baking_QTY.Text) & "^XZ") 'PQ : 발행수량, "ZX : 종료
            '######################## Baking Tray 라벨 발행
            lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-" & TB_Bake_Char.Text & "01" & " / " & tb_After_CC.Text
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & pmic_information & "^FS")
            swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & rcd_information & "^FS")
            swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & TB_Product_Code.Text & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^PQ" & CInt(TB_Bake_QTY.Text) & "^XZ") 'PQ : 발행수량, "ZX : 종료
            '######################## Carrer JIG 라벨 발행
            lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-" & TB_Carrier_Char.Text & "01" & " / " & tb_After_CC.Text
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & pmic_information & "^FS")
            swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & rcd_information & "^FS")
            swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & TB_Product_Code.Text & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^PQ" & CInt(TB_Carrier_QTY.Text) & "^XZ") 'PQ : 발행수량, "ZX : 종료
            '######################## X-Ray Tray 라벨 발행
            lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-" & TB_XRay_Char.Text & "01" & " / " & tb_After_CC.Text
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & pmic_information & "^FS")
            swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & rcd_information & "^FS")
            swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & TB_Product_Code.Text & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^PQ" & CInt(TB_XRay_QTY.Text) & "^XZ") 'PQ : 발행수량, "ZX : 종료
            '######################## X-Ray Box 라벨 발행
            lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-" & TB_Box_Char.Text & "01" & " / " & tb_After_CC.Text
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & pmic_information & "^FS")
            swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & rcd_information & "^FS")
            swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & TB_Product_Code.Text & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^PQ" & CInt(TB_Box_QTY.Text) & "^XZ") 'PQ : 발행수량, "ZX : 종료
            '######################## Reject IC 라벨 발행
            If before_rcd_information = String.Empty Then
                lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-R01" & " / " & tb_After_CC.Text
                swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
                swFile.WriteLine("^MD15") '진하기
                swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
                swFile.WriteLine("^FO000,0065^A0,40,30^FDPMIC : " & before_pmic_information & "^FS")
                swFile.WriteLine("^FO000,0110^A0,40,30^FDRCD : ^FS")
                swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
                swFile.WriteLine("^FO600,0110^A0,40,40^FD" & TB_Qty.Text & "EA" & "^FS")
                swFile.WriteLine("^PQ" & 1 & "^XZ") 'PQ : 발행수량, "ZX : 종료
            Else
                lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-R01" & " / " & tb_After_CC.Text
                swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
                swFile.WriteLine("^MD15") '진하기
                swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
                swFile.WriteLine("^FO000,0065^A0,40,30^FDPMIC : " & before_pmic_information & "^FS")
                swFile.WriteLine("^FO000,0110^A0,40,30^FDRCD : ^FS")
                swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
                swFile.WriteLine("^FO600,0110^A0,40,40^FD" & TB_Qty.Text & "EA" & "^FS")
                swFile.WriteLine("^PQ" & 1 & "^XZ") 'PQ : 발행수량, "ZX : 종료
                lotNo = TB_Lot_No.Text & " / " & TB_YJ_No.Text & "-R02" & " / " & tb_After_CC.Text
                swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
                swFile.WriteLine("^MD15") '진하기
                swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
                swFile.WriteLine("^FO000,0065^A0,40,30^FDPMIC : ^FS")
                swFile.WriteLine("^FO000,0110^A0,40,30^FDRCD : " & before_rcd_information & "^FS")
                swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & Replace(lotNo, " ", "") & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^SF%%%%%%%%%%%%%%%%%%%dd,1^FS")
                swFile.WriteLine("^FO600,0110^A0,40,40^FD" & TB_Qty.Text & "EA" & "^FS")
                swFile.WriteLine("^PQ" & 1 & "^XZ") 'PQ : 발행수량, "ZX : 종료
            End If


            '####### 공정 진행 현황 라벨 발행
            If TB_Option.Text Like "*O" Then '숫자영 영어오
                swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
                swFile.WriteLine("^MD15") '진하기
                swFile.WriteLine("^FO000,0000^A0,60,40^FD* Repair Process Check *^FS")
                swFile.WriteLine("^FO000,0070^A0,35,27^FDPMIC : [  ] Detach > [  ] Auto Line > [  ] Attach^FS")
                swFile.WriteLine("^FO000,0120^A0,35,27^FD        > [  ] Reflow > [  ] AOI^FS")
                'swFile.WriteLine("^FO000,0060^A0,40,30^FDRCD  : Detach > Auto Line > Attach > Reflow > AOI^FS")
                swFile.WriteLine("^PQ1^XZ") 'PQ : 발행수량, "ZX : 종료
            ElseIf TB_Option.Text Like "*Q" Then
                swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
                swFile.WriteLine("^MD15") '진하기
                swFile.WriteLine("^FO000,0000^A0,60,40^FD* Repair Process Check *^FS")
                swFile.WriteLine("^FO000,0070^A0,35,27^FDPMIC : [  ] Detach > [  ] Auto Line > [  ] Attach^FS")
                swFile.WriteLine("^FO000,0120^A0,35,27^FDRCD  : [  ] Detach > [  ] Auto Line > [  ] Reflow > [  ] AOI^FS")
                swFile.WriteLine("^PQ1^XZ") 'PQ : 발행수량, "ZX : 종료
            End If
        ElseIf CB_Parts_Section.Text = "PMIC" Or CB_Parts_Section.Text = "RCD" Then
            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,70,55^FDPart No : " & TB_Product_Code.Text & "^FS")
            swFile.WriteLine("^FO000,0065^A0,50,40^FDLot No. : " & TB_YJ_No.Text & " / " & TB_Lot_No.Text & "^FS")
            swFile.WriteLine("^FO000,0115^A0,50,40^FDTop Marking : " & part_top_marking & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & TB_Product_Code.Text & "/" & TB_Lot_No.Text & "/" & TB_YJ_No.Text & "^FS")
            swFile.WriteLine("^PQ2^XZ") 'PQ : 발행수량, "ZX : 종료
        End If
        swFile.Close()

        '##### 프린터로 전송하는 부분
        Try
            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
            Label_Print_Run = True
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
        End Try

        Return Label_Print_Run

    End Function

    Private Sub YJ_No_Maker()

        Dim yj_no As String = String.Empty

        DBConnect()

        Dim strSQL As String = "call sp_incoming_inspect(null" &
                                            ", null" &
                                            ", null" &
                                            ", 5)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            If Not IsDBNull(sqlDR("yj_no")) Then yj_no = "Y" & sqlDR("yj_no")
        Loop
        sqlDR.Close()

        DBClose()

        If yj_no = String.Empty Then yj_no = "Y00001"

        'yj_no = "Y" & Format(CInt(yj_no.Substring(1, 5)) + 1, "00000")
        TB_YJ_No.Text = yj_no

    End Sub

    Private Sub YJ_No_Maker2(ByVal section As String, ByVal section_char As String)


        Dim yj_no As String = String.Empty

        DBConnect()

        Dim strSQL As String = "call sp_incoming_inspect(null" &
                                            ", '" & section & "'" &
                                            ", null" &
                                            ", 6)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            If Not IsDBNull(sqlDR("yj_no")) Then yj_no = section_char & sqlDR("yj_no")
        Loop
        sqlDR.Close()

        DBClose()

        If yj_no = String.Empty Then yj_no = section_char & "00001"

        'yj_no = section_char & Format(CInt(yj_no.Substring(1, 5)) + 1, "00000")
        TB_YJ_No.Text = yj_no



    End Sub

    Private Sub DB_Write(ByVal write_Select As String)

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try

            Dim qty_check As String = "NG"
            If RB_Qty_OK.Checked Then qty_check = "OK"
            Dim insp_check As String = "NG"
            If RB_Insp_OK.Checked Then insp_check = "OK"
            Dim info_check As String = "NG"
            If RB_Info_OK.Checked Then info_check = "OK"

            Dim write_date As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

            '수입검사 기록
            strSQL = "Insert Into incoming_inspect_result(part_division, lot_no, inspect_date, yj_no"
            strSQL += ", qty_check, inspect_check, information_check, inspect_note, inspector) values"
            strSQL += "('" & CB_Parts_Section.Text & "'"
            strSQL += ",'" & TB_Lot_No.Text & "'"
            strSQL += ",'" & write_date & "'"
            strSQL += ",'" & TB_YJ_No.Text & "'"
            strSQL += ",'" & qty_check & "'"
            strSQL += ",'" & insp_check & "'"
            strSQL += ",'" & info_check & "'"
            strSQL += ",'" & TB_Insp_Memo.Text & "'"
            strSQL += ",'" & loginID & "');"

            If write_Select = "Module" Then
                '라벨 발행 기록
                strSQL += "insert into history_label_print(lot_no, yj_no, process_char"
                strSQL += ", print_qty, printer, print_date, print_method) values("
                strSQL += "'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_YJ_No.Text & "'"
                strSQL += ",'I'"
                strSQL += ",'" & TB_IQC_QTY.Text & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'입고검사');"
                strSQL += "insert into history_label_print(lot_no, yj_no, process_char"
                strSQL += ", print_qty, printer, print_date, print_method) values("
                strSQL += "'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_YJ_No.Text & "'"
                strSQL += ",'B'"
                strSQL += ",'" & TB_Baking_QTY.Text & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'입고검사');"
                strSQL += "insert into history_label_print(lot_no, yj_no, process_char"
                strSQL += ", print_qty, printer, print_date, print_method) values("
                strSQL += "'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_YJ_No.Text & "'"
                strSQL += ",'A'"
                strSQL += ",'" & TB_Bake_QTY.Text & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'입고검사');"
                strSQL += "insert into history_label_print(lot_no, yj_no, process_char"
                strSQL += ", print_qty, printer, print_date, print_method) values("
                strSQL += "'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_YJ_No.Text & "'"
                strSQL += ",'C'"
                strSQL += ",'" & TB_Carrier_QTY.Text & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'입고검사');"
                strSQL += "insert into history_label_print(lot_no, yj_no, process_char"
                strSQL += ", print_qty, printer, print_date, print_method) values("
                strSQL += "'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_YJ_No.Text & "'"
                strSQL += ",'X'"
                strSQL += ",'" & TB_XRay_QTY.Text & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'입고검사');"
                strSQL += "insert into history_label_print(lot_no, yj_no, process_char"
                strSQL += ", print_qty, printer, print_date, print_method) values("
                strSQL += "'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_YJ_No.Text & "'"
                strSQL += ",'O'"
                strSQL += ",'" & TB_Box_QTY.Text & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'입고검사');"
                '########### 폐기자재 회수 라벨
                strSQL += "insert into history_label_print(lot_no, yj_no, process_char"
                strSQL += ", print_qty, printer, print_date, print_method) values("
                strSQL += "'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_YJ_No.Text & "'"
                strSQL += ",'R'"
                If TB_Option.Text Like "*O" Then
                    strSQL += ",'1'"
                ElseIf TB_Option.Text Like "*Q" Then
                    strSQL += ",'2'"
                End If
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'입고검사');"

                'YJ No. 기록
                strSQL += "update basic_lot_information set fab_line = '" & TB_Fab_line.Text & "'"
                strSQL += ", lot_ww = '" & TB_WW.Text & "'"
                strSQL += ", lot_option = '" & TB_Option.Text & "'"
                strSQL += ", lot_status = 'Incoming Inspection Completed'"
                strSQL += ", yj_no = '" & TB_YJ_No.Text & "'"
                strSQL += ", org_product = '" & tb_Before_Product.Text & "'"
                strSQL += " where lot_no = '" & TB_Lot_No.Text & "';"
            ElseIf write_Select = "Material" Then
                'YJ No. 기록
                'strSQL += "update basic_material_information set yj_no = '" & TB_YJ_No.Text & "'"
                'strSQL += " where lot_no = '" & TB_Lot_No.Text & "';"
                strSQL += "insert into basic_material_information(part_division, slip_no, part_no, lot_no"
                strSQL += ", material_qty, material_note, temp_code, first_write_date, writer, write_date, yj_no) values("
                strSQL += "'" & CB_Parts_Section.Text & "'"
                strSQL += ",'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & TB_Product_Code.Text & "'"
                strSQL += ",'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & CInt(TB_Qty.Text) & "'"
                strSQL += ",'" & TB_Insp_Memo.Text & "'"
                strSQL += ",'" & TB_Lot_No.Text & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'" & loginID & "'"
                strSQL += ",'" & write_date & "'"
                strSQL += ",'" & TB_YJ_No.Text & "');"
            End If

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

    End Sub

    Private Sub Inspect_Result_Read()

        thread_LoadingFormStart()

        grid_Inspect_List.Redraw = False
        grid_Inspect_List.Rows.Count = 1

        DBConnect()

        Dim strSQL As String = "call sp_incoming_inspect('" &
            Format(INSP_DATE.Value, "yyyy-MM-dd") &
            "', null, null, 0)"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Dim qty_result As String = "합격"
            If sqlDR("qty_check") = "NG" Then qty_result = "불합격"
            Dim insp_result As String = "합격"
            If sqlDR("inspect_check") = "NG" Then insp_result = "불합격"
            Dim info_result As String = "합격"
            If sqlDR("information_check") = "NG" Then info_result = "불합격"
            Dim insert_string As String = grid_Inspect_List.Rows.Count & vbTab &
                                          Format(sqlDR("inspect_date"), "yyyy-MM-dd HH:mm:ss") & vbTab &
                                          sqlDR("part_division") & vbTab &
                                          sqlDR("product") & vbTab &
                                          sqlDR("lot_no") & vbTab &
                                          sqlDR("chip_qty") & vbTab &
                                          sqlDR("fab_line") & vbTab &
                                          sqlDR("lot_ww") & vbTab &
                                          sqlDR("lot_option") & vbTab &
                                          sqlDR("yj_no") & vbTab &
                                          qty_result & vbTab &
                                          insp_result & vbTab &
                                          info_result & vbTab &
                                          sqlDR("inspector") & vbTab &
                                          sqlDR("inspect_note")
            grid_Inspect_List.AddItem(insert_string)
        Loop
        sqlDR.Close()

        DBClose()

        grid_Inspect_List.AutoSizeCols()
        grid_Inspect_List.Redraw = True

        thread_LoadingFormEnd

    End Sub

    Private Sub TB_Product_Code_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TB_Product_Code.LostFocus

        'If CB_Parts_Section.Text = "Module" Then YJ_No_Maker()

    End Sub

    Private Sub TB_Lot_No_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Lot_No.KeyDown

        If e.KeyCode = 13 And Not TB_Lot_No.Text = String.Empty Then
            If Strings.Left(TB_Lot_No.Text, 1) = "+" Then
                '+000020000023 TCKM35.2
                '자재일 경우
                Dim org_lotno As String = TB_Lot_No.Text
                TB_Lot_No.Text = org_lotno.Split(" ")(0).Substring(9, 4) & org_lotno.Split(" ")(1)
                TB_Qty.Text = Format(CDbl(org_lotno.Split(" ")(0).Substring(1, 8)), "#,##0")
            End If

            If Lot_Exist_Check() = True Then
                MsgBox("수입검사 완료된 Lot No.입니다.", MsgBoxStyle.Information, form_Msgbox_String)
                TB_Lot_No.SelectAll()
                TB_Lot_No.Focus()
                TB_Qty.Text = String.Empty
                Exit Sub
            End If

            If CB_Parts_Section.Text = "Module" Or CB_Parts_Section.Text = "Component" Then
                If LOT_Info_Load(CB_Parts_Section.Text) = False Then
                    MsgBox("입고 등록되지 않은 Lot No. 입니다." & vbCrLf &
                       "확인하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
                    TB_Lot_No.SelectAll()
                    TB_Qty.Text = String.Empty
                End If
            End If

            If CB_Parts_Section.Text = "Module" Then
                TB_Qty.Focus()
                TB_Qty.SelectAll()
            ElseIf CB_Parts_Section.Text = "PMIC" Or CB_Parts_Section.Text = "RCD" Then
                'Module외 자재들은 바로 입고등록 진행 하면 된다.
                If TB_Qty.Text = String.Empty Then
                    TB_Qty.SelectAll()
                    TB_Qty.Focus()
                Else
                    BTN_SavePrint_Click(Nothing, Nothing)
                End If
            End If
        End If

    End Sub

    Private Sub TB_Lot_No_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TB_Lot_No.LostFocus

        'If Lot_Exist_Check() = True Then
        '    MsgBox("수입검사 완료된 Lot No.입니다.", MsgBoxStyle.Information, form_Msgbox_String)
        '    Exit Sub
        'End If
        'LOT_Info_Load()

    End Sub

    Private Function LOT_Info_Load(ByVal part_section As String) As Boolean

        Dim existLotno As Boolean = False
        c_product_code = String.Empty
        c_qty = 0
        c_fab_line = String.Empty
        c_ww = String.Empty
        c_option = String.Empty

        DBConnect()

        Dim strSQL As String = "call sp_incoming_inspect(null" &
                                                ",'" & TB_Lot_No.Text & "'" &
                                                ",null , 1)"
        If part_section = "PMIC" Or part_section = "RCD" Then
            strSQL = "call sp_incoming_inspect(null" &
                ",'" & TB_Lot_No.Text & "'" &
                ",null, 4)"
        End If

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            If part_section = "Module" Then
                c_product_code = sqlDR("product")
                c_qty = sqlDR("chip_qty")
                c_fab_line = sqlDR("fab_line")
                c_ww = sqlDR("lot_ww")
                c_option = sqlDR("lot_option")
            ElseIf part_section = "PMIC" Or part_section = "RCD" Then
                c_qty = sqlDR("material_qty")
                c_product_code = sqlDR("part_no")
            End If
        Loop
        sqlDR.Close()

        DBClose()

        If c_qty = 0 Then
            existLotno = False
        Else
            existLotno = True
        End If

        Return existLotno

    End Function

    Private Function Lot_Exist_Check() As Boolean

        Lot_Exist_Check = False

        DBConnect()

        Dim strSQL As String = "select lot_no from incoming_inspect_result"
        strSQL += " where lot_no = '" & TB_Lot_No.Text & "'"

        Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
        Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

        Do While sqlDR.Read
            Lot_Exist_Check = True
        Loop
        sqlDR.Close()

        DBClose()

        Return Lot_Exist_Check

    End Function

    Private Sub INSP_DATE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles INSP_DATE.ValueChanged

        Label17.Text = Format(INSP_DATE.Value, "yyyy년 MM월 dd일 입고이력")
        Inspect_Result_Read()

    End Sub

    Private Sub CB_Parts_Section_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Parts_Section.SelectedIndexChanged

        If Not CB_Parts_Section.Text = String.Empty Then
            TB_Product_Code.Focus()
            TB_Product_Code.SelectAll()
        End If

    End Sub

    Private Sub TB_Product_Code_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Product_Code.KeyDown

        If e.KeyCode = 13 And Not TB_Product_Code.Text = String.Empty Then

            If TB_Product_Code.Text Like "M*" Then
                CB_Parts_Section.Text = "Module"
            ElseIf TB_Product_Code.Text Like "K*" Then
                CB_Parts_Section.Text = "Component"
            Else
                Dim a As String = Strings.Right(TB_Product_Code.Text, 1)
                Dim b As String = Strings.Left(TB_Product_Code.Text, 1)
                If a = "+" Then
                    If sub_Part_Information("pmic_partno", Replace(TB_Product_Code.Text.Split(" ")(1), "+", String.Empty)) > 0 Then
                        CB_Parts_Section.Text = "PMIC"
                        TB_Product_Code.Text = Replace(TB_Product_Code.Text.Split(" ")(1), "+", String.Empty)
                    ElseIf sub_Part_Information("rcd_partno", Replace(TB_Product_Code.Text.Split(" ")(1), "+", String.Empty)) > 0 Then
                        CB_Parts_Section.Text = "RCD"
                        TB_Product_Code.Text = Replace(TB_Product_Code.Text.Split(" ")(1), "+", String.Empty)
                    Else
                        MsgBox("자재 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                        TB_Product_Code.SelectAll()
                    End If
                ElseIf b = "+" Then
                    MsgBox("Part No. Barcode를 스캔하여 주십시오.", MsgBoxStyle.Information, form_Msgbox_String)
                Else
                    If sub_Part_Information("pmic_partno", TB_Product_Code.Text) > 0 Then
                        CB_Parts_Section.Text = "PMIC"
                    ElseIf sub_Part_Information("rcd_partno", TB_Product_Code.Text) > 0 Then
                        CB_Parts_Section.Text = "RCD"
                    Else
                        MsgBox("자재 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                        TB_Product_Code.SelectAll()
                    End If
                End If
            End If

            If CB_Parts_Section.Text = "Module" Then
                Dim ccInformation() As String = cc_Information(TB_Product_Code.Text).Split("/")

                If ccInformation.Length > 1 Then
                    tb_After_CC.Text = TB_Product_Code.Text.Substring(16, 2)
                    tb_After_PMIC.Text = ccInformation(0)
                    change_pmic_top_marking = ccInformation(1)
                    tb_After_RCD.Text = ccInformation(2)
                    change_rcd_top_marking = ccInformation(3)
                    TB_Lot_No.Focus()
                    TB_Lot_No.SelectAll()
                    YJ_No_Maker()
                Else
                    TB_Product_Code.SelectAll()
                    MsgBox("Customer Code에 대한 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
                End If
            ElseIf CB_Parts_Section.Text = "PMIC" Or CB_Parts_Section.Text = "RCD" Then
                YJ_No_Maker2(CB_Parts_Section.Text, CB_Parts_Section.Text.Substring(0, 1))
                TB_Lot_No.Focus()
                TB_Lot_No.SelectAll()
            End If
        End If

    End Sub

    Private Function sub_Part_Information(ByVal section As String, ByVal part_no As String) As Integer

        Dim total_count As Integer = 0

        DBConnect()

        Try
            Dim strSQL As String = "call sp_incoming_inspect(null, '" & section & "', '" & part_no & "', 3)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                total_count += 1
                part_top_marking = sqlDR("top_marking")
            Loop
            sqlDR.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try
        DBClose()

        Return total_count

    End Function

    Private Function cc_Information(ByVal product_code As String) As String

        cc_Information = String.Empty

        DBConnect()

        Try
            Dim strSQL As String = "call sp_incoming_inspect(null, '" & product_code.Substring(16, 2) & "', null, 2)"

            Dim sqlCmd As New MySqlCommand(strSQL, dbConnection1)
            Dim sqlDR As MySqlDataReader = sqlCmd.ExecuteReader

            Do While sqlDR.Read
                cc_Information = sqlDR("pmic_partno") & "/" &
                    sqlDR("pmic_top_marking") & "/" &
                    sqlDR("rcd_partno") & "/" &
                    sqlDR("rcd_top_marking")
            Loop
            sqlDR.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
        End Try
        DBClose()

        Return cc_Information

    End Function

    Private Sub TB_Qty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Qty.KeyDown

        Select Case e.KeyCode
            Case 13
            Case 8
            Case 96 To 105, 48 To 57
            Case Else
                e.SuppressKeyPress = True
        End Select

        If e.KeyCode = 13 And Not TB_Qty.Text = String.Empty Then
            If CB_Parts_Section.Text = "Module" Then
                Print_Qty_Cal()
                TB_Fab_line.Focus()
                TB_Fab_line.SelectAll()
            ElseIf CB_Parts_Section.Text = "PMIC" Or CB_Parts_Section.Text = "RCD" Then
                BTN_SavePrint_Click(Nothing, Nothing)
            End If
        End If

    End Sub

    Private Sub Print_Qty_Cal()

        Dim label_plus As Integer = 0

        If CheckBox1.Checked Then label_plus = 3

        If CB_Parts_Section.Text = "Module" Then
            TB_IQC_QTY.Text = Math.Ceiling(CInt(TB_Qty.Text) / CInt(TB_IQC_Max.Text))
            TB_Baking_QTY.Text = Math.Ceiling(CInt(TB_Qty.Text) / CInt(TB_Baking_Max.Text)) + label_plus
            TB_Bake_QTY.Text = Math.Ceiling(CInt(TB_Qty.Text) / CInt(TB_Bake_Max.Text)) + label_plus
            TB_Carrier_QTY.Text = Math.Ceiling(CInt(TB_Qty.Text) / CInt(TB_Carrier_Max.Text)) + label_plus
            TB_XRay_QTY.Text = 1
            TB_Box_QTY.Text = 1
        Else
            TB_IQC_QTY.Text = 0
            TB_Baking_QTY.Text = 0
            TB_Bake_QTY.Text = 0
            TB_Carrier_QTY.Text = 0
            TB_XRay_QTY.Text = 0
            TB_Box_QTY.Text = 0
        End If

    End Sub

    Private Sub TB_Fab_line_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Fab_line.KeyDown

        If e.KeyCode = 13 And Not TB_Fab_line.Text = String.Empty Then
            TB_WW.Focus()
            TB_WW.SelectAll()
        ElseIf e.KeyCode = 13 And Not CB_Parts_Section.Text = "Module" Then
            TB_WW.Focus()
            TB_WW.SelectAll()
        End If

    End Sub

    Private Sub TB_WW_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_WW.KeyDown

        If e.KeyCode = 13 And Not TB_WW.Text = String.Empty Then
            TB_Option.Focus()
            TB_Option.SelectAll()
        ElseIf e.KeyCode = 13 And Not CB_Parts_Section.Text = "Module" Then
            TB_Option.Focus()
            TB_Option.SelectAll()
        End If

    End Sub

    Private Sub TB_Option_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TB_Option.KeyDown

        If e.KeyCode = 13 And Not TB_Option.Text = String.Empty Then
            tb_Before_Product.Focus()
            tb_Before_Product.SelectAll()
        ElseIf e.KeyCode = 13 And Not CB_Parts_Section.Text = "Module" Then
            tb_Before_Product.Focus()
            tb_Before_Product.SelectAll()
        End If

    End Sub

    Private Sub Form_CLose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form_CLose.Click

        Me.Dispose()

    End Sub

    Private Sub BTN_Setting_Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Setting_Open.Click

        If SplitContainer1.Panel1Collapsed = True Then
            SplitContainer1.Panel1Collapsed = False
            BTN_Setting_Open.Text = "프린터 설정 닫기"
        Else
            SplitContainer1.Panel1Collapsed = True
            BTN_Setting_Open.Text = "프린터 설정 열기"
        End If

    End Sub

    Private Sub TB_IQC_Max_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_IQC_Max.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Qty", "IQC", TB_IQC_Max.Text)

    End Sub

    Private Sub TB_Bake_Max_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bake_Max.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Qty", "Bake", TB_Bake_Max.Text)

    End Sub

    Private Sub TB_Scope_Max_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_XRay_Max.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Qty", "X-Ray", TB_XRay_Max.Text)

    End Sub

    Private Sub TB_Carrier_Max_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Carrier_Max.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Qty", "Carrier", TB_Carrier_Max.Text)

    End Sub

    Private Sub TB_Baking_Max_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Baking_Max.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Qty", "Baking", TB_Baking_Max.Text)

    End Sub

    Private Sub TB_BGA_Tray_Max_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Box_Max.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Qty", "Box", TB_Box_Max.Text)

    End Sub

    Private Sub TB_Bake_Char_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Bake_Char.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Char", "Bake", TB_Bake_Char.Text)

    End Sub

    Private Sub TB_IQC_Char_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_IQC_Char.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Char", "IQC", TB_IQC_Char.Text)

    End Sub

    Private Sub TB_Scope_Char_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_XRay_Char.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Char", "X-Ray", TB_XRay_Char.Text)

    End Sub

    Private Sub TB_Carrier_Char_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Carrier_Char.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Char", "Carrier", TB_Carrier_Char.Text)

    End Sub

    Private Sub TB_Baking_Char_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Baking_Char.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Char", "Baking", TB_Baking_Char.Text)

    End Sub

    Private Sub TB_Box_Char_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Box_Char.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\Char", "Box", TB_Box_Char.Text)

    End Sub

    Private Sub CB_Cable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_Cable.SelectedIndexChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "COM/LPT", cb_Cable.Text)

        If cb_Cable.Text = "USB" Then
            PrinterListLoad()
        End If

    End Sub

    Private Sub TB_Port_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Port.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Port", TB_Port.Text)

    End Sub

    Private Sub TB_Width_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_LEFT_Loc.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Width", TB_LEFT_Loc.Text)

    End Sub

    Private Sub TB_Height_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_TOP_Loc.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Height", TB_TOP_Loc.Text)

    End Sub

    Private Sub TB_BaudRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_BaudRate.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "BaudRate", TB_BaudRate.Text)

    End Sub

    Private Sub TB_DataBits_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_DataBits.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "DataBits", TB_DataBits.Text)

    End Sub

    Private Sub TB_Parity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Parity.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Parity", TB_Parity.Text)

    End Sub

    Private Sub TB_StopBits_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_StopBits.TextChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "StopBits", TB_StopBits.Text)

    End Sub

    Private Sub Setting_LOAD()

        cb_Cable.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "COM/LPT", "COM")
        cb_PrinterList.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Printer Name", "")
        TB_Port.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Port", 1)
        TB_LEFT_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Width", 0)
        TB_TOP_Loc.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Height", 0)
        TB_BaudRate.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "BaudRate", 9600)
        TB_DataBits.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "DataBits", 8)
        TB_Parity.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Parity", 0)
        TB_StopBits.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "StopBits", 1)

        TB_IQC_Max.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Qty", "IQC", 400)
        TB_Bake_Max.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Qty", "Bake", 25)
        TB_XRay_Max.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Qty", "X-Ray", 25)
        TB_Carrier_Max.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Qty", "Carrier", 12)
        TB_Baking_Max.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Qty", "Baking", 25)
        TB_Box_Max.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Qty", "Box", 384)

        TB_IQC_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "IQC", "I")
        TB_Bake_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Bake", "A")
        TB_XRay_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "X-Ray", "X")
        TB_Carrier_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Carrier", "C")
        TB_Baking_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Baking", "B")
        TB_Box_Char.Text = registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\Char", "Box", "O")

        If registryEdit.ReadRegKey("Software\Yujin\Repair System\Label\PRINTER", "Print Use", True) = True Then
            CB_Print_Use.Checked = True
        Else
            CB_Print_Use.Checked = False
        End If

    End Sub

    Private Sub CB_Print_Use_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Print_Use.CheckedChanged

        If CB_Print_Use.Checked Then
            registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Print Use", True)
        Else
            registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Print Use", False)
        End If

    End Sub


    Private Sub PrinterListLoad()

        cb_PrinterList.Items.Clear()

        For i As Integer = 0 To Printing.PrinterSettings.InstalledPrinters.Count - 1
            Dim Printers As String = Printing.PrinterSettings.InstalledPrinters.Item(i)
            cb_PrinterList.Items.Add(Printers)
        Next

        cb_PrinterList.Sorted = True

    End Sub

    Private Sub cb_PrinterList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_PrinterList.SelectedIndexChanged

        registryEdit.WriteRegKey("Software\Yujin\Repair System\Label\PRINTER", "Printer Name", cb_PrinterList.Text)

    End Sub

    Private Sub tb_Change_Product_KeyDown(sender As Object, e As KeyEventArgs) Handles tb_Before_Product.KeyDown

        If e.KeyCode = 13 And Not tb_Before_Product.Text = String.Empty Then
            Dim ccInformation() As String = cc_Information(tb_Before_Product.Text).Split("/")

            If ccInformation.Length > 1 Then
                tb_before_CC.Text = tb_Before_Product.Text.Substring(16, 2)
                tb_Before_PMIC.Text = ccInformation(0)
                before_pmic_top_marking = ccInformation(1)
                tb_Before_RCD.Text = ccInformation(2)
                before_rcd_top_marking = ccInformation(3)
                BTN_SavePrint_Click(Nothing, Nothing)
            Else
                tb_Before_Product.SelectAll()
                MsgBox("Customer Code에 대한 정보를 찾을 수 없습니다.", MsgBoxStyle.Information, form_Msgbox_String)
            End If
        End If

    End Sub

    Private Sub btn_TestLabel_Click(sender As Object, e As EventArgs) Handles btn_TestLabel.Click

        Dim test_result As String = print_test()

        If Not test_result = "Success" Then
            MsgBox("프린트에 실패 하였습니다." & vbCrLf & test_result, MsgBoxStyle.Critical, form_Msgbox_String)
        Else
            MsgBox("테스트 프린트 완료.", MsgBoxStyle.Information, form_Msgbox_String)
        End If

    End Sub

    Private Function print_test() As String

        print_test = String.Empty

        '##### 프린터로 전송하는 부분
        Try
            Dim swFile As StreamWriter =
            New StreamWriter(Application.StartupPath & "\print.txt", True, System.Text.Encoding.GetEncoding(949))

            Dim lotNo As String = "SSLOTNOT" & " / " & "YJLOTNO" & "-" & "T" & "01" & " / " & "ZZ"
            swFile.WriteLine("^XZ~JA^XZ")
            swFile.WriteLine("^XA^LH" & TB_LEFT_Loc.Text & ",0^LT" & TB_TOP_Loc.Text) 'LH : 가로위치, LT : 세로위치
            swFile.WriteLine("^MD15") '진하기
            swFile.WriteLine("^FO000,0000^A0,60,40^FD" & lotNo & "^SF%%%%%%%%%%%%%%%%%%%%%dd,1^FS")
            swFile.WriteLine("^FO000,0060^A0,40,30^FDPMIC : " & "PMIC PARTNO ( Marking )" & "^FS")
            swFile.WriteLine("^FO000,0095^A0,40,30^FDRCD : " & "RCD PARTNO ( Marking )" & "^FS")
            'M321R8GA0BB0-CQKZJ-WGB
            swFile.WriteLine("^FO000,0135^A0,35,25^FDP/N : " & "M321R8GA0BB0-CQKZJ-WGB" & "^FS")
            swFile.WriteLine("^FO600,0000^BXN,3,200^FD" & lotNo & "/" & tb_After_PMIC.Text & "/" & tb_After_RCD.Text & "^FS")
            swFile.WriteLine("^FO600,0110^A0,40,40^FD" & "250EA" & "^FS")
            swFile.WriteLine("^PQ1^XZ") 'PQ : 발행수량, "ZX : 종료
            swFile.Close()

            If cb_Cable.Text = "LPT" Then
                File.Copy(Application.StartupPath & "\print.txt", "LPT" & TB_Port.Text)
            ElseIf cb_Cable.Text = "COM" Then
                ComPort.PortName = "COM" & TB_Port.Text
                ComPort.BaudRate = TB_BaudRate.Text
                ComPort.Parity = TB_Parity.Text
                ComPort.DataBits = TB_DataBits.Text
                ComPort.StopBits = TB_StopBits.Text
                ComPort.Encoding = System.Text.Encoding.Default
                'ComPort.Handshake = 

                Call ComPort.Open()
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    ComPort.WriteLine(str)
                End While

                sr.Close()
                ComPort.Close()
            ElseIf cb_Cable.Text = "USB" Then
                Dim p As New PrintRaw
                Dim s As New StringBuilder
                Dim fs As System.IO.FileStream
                Dim sr As System.IO.StreamReader
                fs = System.IO.File.Open(Application.StartupPath & "\print.txt", IO.FileMode.Open) ' 파일 열기
                sr = New System.IO.StreamReader(fs) ' 스트림리더에 연결
                Dim str As String = String.Empty

                While sr.Peek() >= 0
                    str = sr.ReadLine() ' 한줄씩 읽기
                    s.AppendLine(str)
                End While

                sr.Close()
                p.Print(s, cb_PrinterList.Text)
            End If
            File.Delete(Application.StartupPath & "\print.txt")
            print_test = "Success"
        Catch ex As Exception
            File.Delete(Application.StartupPath & "\print.txt")
            print_test = ex.Message
        End Try

        Return print_test

    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If Not TB_Qty.Text.Equals(String.Empty) Then Print_Qty_Cal()

    End Sub

    Private Sub btn_SaveCancel_Click(sender As Object, e As EventArgs) Handles btn_SaveCancel.Click

        Dim sel_Row As Integer = grid_Inspect_List.Row
        Dim sel_YJ_No As String = grid_Inspect_List(sel_Row, 9)

        If MsgBox("선택된 YJ No. : " & sel_YJ_No & " 를(을) 취소 하시겠습니까?",
                  MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                  form_Msgbox_String) = MsgBoxResult.No Then Exit Sub

        DBConnect()

        Dim sqlTran As MySqlTransaction
        Dim sqlCmd As MySqlCommand
        Dim strSQL As String = String.Empty

        sqlTran = dbConnection1.BeginTransaction

        Try
            strSQL += "delete from incoming_inspect_result"
            strSQL += " where yj_no = '" & sel_YJ_No & "';"

            strSQL += "update basic_lot_information set lot_status = 'Ready', yj_no = null"
            strSQL += " where yj_no = '" & sel_YJ_No & "';"

            strSQL += "delete from history_label_print"
            strSQL += " where yj_no = '" & sel_YJ_No & "';"

            If Not strSQL = String.Empty Then
                sqlCmd = New MySqlCommand(strSQL, dbConnection1)
                sqlCmd.Transaction = sqlTran
                sqlCmd.ExecuteNonQuery()

                sqlTran.Commit()
            End If
        Catch ex As MySqlException
            sqlTran.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, form_Msgbox_String)
            Exit Sub
        End Try

        DBClose()

        MsgBox("취소완료.", MsgBoxStyle.Information, form_Msgbox_String)

        Inspect_Result_Read()

    End Sub

    Private Sub TB_Product_Code_TextChanged(sender As Object, e As EventArgs) Handles TB_Product_Code.TextChanged

    End Sub

    Private Sub TB_Lot_No_TextChanged(sender As Object, e As EventArgs) Handles TB_Lot_No.TextChanged

    End Sub

    Private Sub tb_Before_Product_TextChanged(sender As Object, e As EventArgs) Handles tb_Before_Product.TextChanged

    End Sub

    Private Sub grid_Inspect_List_MouseClick(sender As Object, e As MouseEventArgs) Handles grid_Inspect_List.MouseClick

        If e.Button = MouseButtons.Right And grid_Inspect_List.MouseRow > 0 Then
            grid_Inspect_List.Row = grid_Inspect_List.MouseRow
            If loginID = "ADMIN" Then
                grid_Menu.Show(grid_Inspect_List, New Point(e.X, e.Y))
            End If
        End If

    End Sub

    Private Sub grid_Inspect_List_MouseMove(sender As Object, e As MouseEventArgs) Handles grid_Inspect_List.MouseMove

        If sender.MouseRow > 0 And
            e.Button = MouseButtons.Left Then

            md_ETC.cellCal(sender, Nothing)
        End If

    End Sub

    Private Sub grid_LotList_LostFocus(sender As Object, e As EventArgs) Handles grid_Inspect_List.LostFocus

        frm_Main.lb_Status.Text = String.Empty

    End Sub

    Private Sub grid_Inspect_List_KeyDown(sender As Object, e As KeyEventArgs) Handles grid_Inspect_List.KeyDown

        If e.Shift Then
            md_ETC.cellCal(sender, Nothing)
        Else
            frm_Main.lb_Status.Text = String.Empty
        End If

    End Sub
End Class