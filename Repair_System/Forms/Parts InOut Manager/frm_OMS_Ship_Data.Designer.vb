<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_OMS_Ship_Data
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_OMS_Ship_Data))
        Me.tb_Area_Flag = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.tb_To_Site = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.tb_Assembly_Site = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.tb_TPN_No = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.tb_Match_Field = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.tb_To_Area = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.tb_From_Area = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.tb_Check_Option = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.tb_Customer_Name = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.tb_Process_ID = New System.Windows.Forms.TextBox()
        Me.tb_PCB_Name = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tb_Start_Tag = New System.Windows.Forms.TextBox()
        Me.tb_Work_Week = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tb_Company_Code = New System.Windows.Forms.TextBox()
        Me.tb_Loss_Qty = New System.Windows.Forms.TextBox()
        Me.GRID_Lot_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.tb_Inkless = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tb_Mold_DateTime = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tb_MDL_LOC_ID = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btn_Search = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btn_Data_Output = New System.Windows.Forms.ToolStripButton()
        Me.Form_CLose = New System.Windows.Forms.ToolStripButton()
        Me.btn_Table_Show = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GRID_Data_List = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        CType(Me.GRID_Lot_List, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.GRID_Data_List, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SuspendLayout()
        '
        'tb_Area_Flag
        '
        Me.tb_Area_Flag.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Area_Flag.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Area_Flag.Location = New System.Drawing.Point(978, 62)
        Me.tb_Area_Flag.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Area_Flag.Name = "tb_Area_Flag"
        Me.tb_Area_Flag.Size = New System.Drawing.Size(68, 21)
        Me.tb_Area_Flag.TabIndex = 38
        Me.tb_Area_Flag.Text = "TEST"
        '
        'Label21
        '
        Me.Label21.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label21.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label21.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.White
        Me.Label21.Location = New System.Drawing.Point(845, 62)
        Me.Label21.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(133, 21)
        Me.Label21.TabIndex = 31
        Me.Label21.Text = "Area Flag"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_To_Site
        '
        Me.tb_To_Site.BackColor = System.Drawing.SystemColors.Window
        Me.tb_To_Site.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_To_Site.Location = New System.Drawing.Point(978, 39)
        Me.tb_To_Site.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_To_Site.Name = "tb_To_Site"
        Me.tb_To_Site.Size = New System.Drawing.Size(68, 21)
        Me.tb_To_Site.TabIndex = 36
        Me.tb_To_Site.Text = "S"
        '
        'Label23
        '
        Me.Label23.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label23.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label23.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.White
        Me.Label23.Location = New System.Drawing.Point(845, 39)
        Me.Label23.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(133, 21)
        Me.Label23.TabIndex = 29
        Me.Label23.Text = "To Site"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Assembly_Site
        '
        Me.tb_Assembly_Site.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Assembly_Site.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Assembly_Site.Location = New System.Drawing.Point(978, 16)
        Me.tb_Assembly_Site.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Assembly_Site.Name = "tb_Assembly_Site"
        Me.tb_Assembly_Site.Size = New System.Drawing.Size(68, 21)
        Me.tb_Assembly_Site.TabIndex = 34
        Me.tb_Assembly_Site.Text = "UJ"
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label24.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label24.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.White
        Me.Label24.Location = New System.Drawing.Point(845, 16)
        Me.Label24.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(133, 21)
        Me.Label24.TabIndex = 27
        Me.Label24.Text = "Assembly Site"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_TPN_No
        '
        Me.tb_TPN_No.BackColor = System.Drawing.SystemColors.Window
        Me.tb_TPN_No.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_TPN_No.Location = New System.Drawing.Point(777, 85)
        Me.tb_TPN_No.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_TPN_No.Name = "tb_TPN_No"
        Me.tb_TPN_No.Size = New System.Drawing.Size(68, 21)
        Me.tb_TPN_No.TabIndex = 32
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.White
        Me.Label17.Location = New System.Drawing.Point(644, 85)
        Me.Label17.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(133, 21)
        Me.Label17.TabIndex = 25
        Me.Label17.Text = "TPN No"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Match_Field
        '
        Me.tb_Match_Field.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Match_Field.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Match_Field.Location = New System.Drawing.Point(777, 62)
        Me.tb_Match_Field.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Match_Field.Name = "tb_Match_Field"
        Me.tb_Match_Field.Size = New System.Drawing.Size(68, 21)
        Me.tb_Match_Field.TabIndex = 30
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label18.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.White
        Me.Label18.Location = New System.Drawing.Point(644, 62)
        Me.Label18.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(133, 21)
        Me.Label18.TabIndex = 23
        Me.Label18.Text = "Match Field"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_To_Area
        '
        Me.tb_To_Area.BackColor = System.Drawing.SystemColors.Window
        Me.tb_To_Area.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_To_Area.Location = New System.Drawing.Point(777, 39)
        Me.tb_To_Area.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_To_Area.Name = "tb_To_Area"
        Me.tb_To_Area.Size = New System.Drawing.Size(68, 21)
        Me.tb_To_Area.TabIndex = 28
        Me.tb_To_Area.Text = "WMDL"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.White
        Me.Label19.Location = New System.Drawing.Point(644, 39)
        Me.Label19.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(133, 21)
        Me.Label19.TabIndex = 21
        Me.Label19.Text = "To Area"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_From_Area
        '
        Me.tb_From_Area.BackColor = System.Drawing.SystemColors.Window
        Me.tb_From_Area.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_From_Area.Location = New System.Drawing.Point(777, 16)
        Me.tb_From_Area.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_From_Area.Name = "tb_From_Area"
        Me.tb_From_Area.Size = New System.Drawing.Size(68, 21)
        Me.tb_From_Area.TabIndex = 26
        Me.tb_From_Area.Text = "JTST"
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label20.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.White
        Me.Label20.Location = New System.Drawing.Point(644, 16)
        Me.Label20.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(133, 21)
        Me.Label20.TabIndex = 19
        Me.Label20.Text = "From Area"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Check_Option
        '
        Me.tb_Check_Option.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Check_Option.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Check_Option.Location = New System.Drawing.Point(576, 85)
        Me.tb_Check_Option.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Check_Option.Name = "tb_Check_Option"
        Me.tb_Check_Option.Size = New System.Drawing.Size(68, 21)
        Me.tb_Check_Option.TabIndex = 24
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label16.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.White
        Me.Label16.Location = New System.Drawing.Point(443, 85)
        Me.Label16.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(133, 21)
        Me.Label16.TabIndex = 17
        Me.Label16.Text = "Check Option"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(443, 62)
        Me.Label15.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(133, 21)
        Me.Label15.TabIndex = 15
        Me.Label15.Text = "PCB Vendor"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Customer_Name
        '
        Me.tb_Customer_Name.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Customer_Name.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Customer_Name.Location = New System.Drawing.Point(576, 39)
        Me.tb_Customer_Name.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Customer_Name.Name = "tb_Customer_Name"
        Me.tb_Customer_Name.Size = New System.Drawing.Size(68, 21)
        Me.tb_Customer_Name.TabIndex = 20
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label14.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(443, 39)
        Me.Label14.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(133, 21)
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "Customer Name"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Process_ID
        '
        Me.tb_Process_ID.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Process_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Process_ID.Location = New System.Drawing.Point(576, 16)
        Me.tb_Process_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Process_ID.Name = "tb_Process_ID"
        Me.tb_Process_ID.Size = New System.Drawing.Size(68, 21)
        Me.tb_Process_ID.TabIndex = 18
        '
        'tb_PCB_Name
        '
        Me.tb_PCB_Name.BackColor = System.Drawing.SystemColors.Window
        Me.tb_PCB_Name.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_PCB_Name.Location = New System.Drawing.Point(576, 62)
        Me.tb_PCB_Name.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_PCB_Name.Name = "tb_PCB_Name"
        Me.tb_PCB_Name.Size = New System.Drawing.Size(68, 21)
        Me.tb_PCB_Name.TabIndex = 22
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.White
        Me.Label13.Location = New System.Drawing.Point(443, 16)
        Me.Label13.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(133, 21)
        Me.Label13.TabIndex = 11
        Me.Label13.Text = "Process ID"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(242, 16)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(133, 21)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Start Tag"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Start_Tag
        '
        Me.tb_Start_Tag.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Start_Tag.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Start_Tag.Location = New System.Drawing.Point(375, 16)
        Me.tb_Start_Tag.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Start_Tag.Name = "tb_Start_Tag"
        Me.tb_Start_Tag.Size = New System.Drawing.Size(68, 21)
        Me.tb_Start_Tag.TabIndex = 10
        Me.tb_Start_Tag.Text = "LS1"
        '
        'tb_Work_Week
        '
        Me.tb_Work_Week.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Work_Week.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Work_Week.Location = New System.Drawing.Point(375, 85)
        Me.tb_Work_Week.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Work_Week.Name = "tb_Work_Week"
        Me.tb_Work_Week.Size = New System.Drawing.Size(68, 21)
        Me.tb_Work_Week.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(242, 39)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(133, 21)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Company Code"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(242, 85)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(133, 21)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Work Week"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Company_Code
        '
        Me.tb_Company_Code.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Company_Code.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Company_Code.Location = New System.Drawing.Point(375, 39)
        Me.tb_Company_Code.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Company_Code.Name = "tb_Company_Code"
        Me.tb_Company_Code.Size = New System.Drawing.Size(68, 21)
        Me.tb_Company_Code.TabIndex = 12
        Me.tb_Company_Code.Text = "UJ"
        '
        'tb_Loss_Qty
        '
        Me.tb_Loss_Qty.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Loss_Qty.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Loss_Qty.Location = New System.Drawing.Point(375, 62)
        Me.tb_Loss_Qty.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Loss_Qty.Name = "tb_Loss_Qty"
        Me.tb_Loss_Qty.Size = New System.Drawing.Size(68, 21)
        Me.tb_Loss_Qty.TabIndex = 14
        Me.tb_Loss_Qty.Text = "0"
        '
        'GRID_Lot_List
        '
        Me.GRID_Lot_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_Lot_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_Lot_List.Location = New System.Drawing.Point(0, 0)
        Me.GRID_Lot_List.Name = "GRID_Lot_List"
        Me.GRID_Lot_List.Rows.Count = 2
        Me.GRID_Lot_List.Rows.DefaultSize = 20
        Me.GRID_Lot_List.Size = New System.Drawing.Size(1264, 617)
        Me.GRID_Lot_List.StyleInfo = resources.GetString("GRID_Lot_List.StyleInfo")
        Me.GRID_Lot_List.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.TextBox4)
        Me.Panel1.Controls.Add(Me.tb_Inkless)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.tb_Mold_DateTime)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.tb_MDL_LOC_ID)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.tb_Area_Flag)
        Me.Panel1.Controls.Add(Me.Label21)
        Me.Panel1.Controls.Add(Me.tb_To_Site)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.tb_Assembly_Site)
        Me.Panel1.Controls.Add(Me.Label24)
        Me.Panel1.Controls.Add(Me.tb_TPN_No)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.tb_Match_Field)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.tb_To_Area)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.tb_From_Area)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.tb_Check_Option)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.tb_PCB_Name)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.tb_Customer_Name)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.tb_Process_ID)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.tb_Start_Tag)
        Me.Panel1.Controls.Add(Me.tb_Work_Week)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.tb_Company_Code)
        Me.Panel1.Controls.Add(Me.tb_Loss_Qty)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1264, 127)
        Me.Panel1.TabIndex = 0
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(5, 111)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(99, 12)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "총 출고 Module :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 94)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(74, 12)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "총 출고 Lot :"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label8.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(7, 16)
        Me.Label8.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(133, 21)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Slip NO."
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox4
        '
        Me.TextBox4.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox4.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.TextBox4.Location = New System.Drawing.Point(140, 16)
        Me.TextBox4.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(90, 21)
        Me.TextBox4.TabIndex = 1
        '
        'tb_Inkless
        '
        Me.tb_Inkless.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Inkless.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Inkless.Location = New System.Drawing.Point(1179, 39)
        Me.tb_Inkless.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Inkless.Name = "tb_Inkless"
        Me.tb_Inkless.Size = New System.Drawing.Size(68, 21)
        Me.tb_Inkless.TabIndex = 42
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(1046, 39)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(133, 21)
        Me.Label7.TabIndex = 37
        Me.Label7.Text = "Inkless"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_Mold_DateTime
        '
        Me.tb_Mold_DateTime.BackColor = System.Drawing.SystemColors.Window
        Me.tb_Mold_DateTime.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_Mold_DateTime.Location = New System.Drawing.Point(1179, 16)
        Me.tb_Mold_DateTime.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_Mold_DateTime.Name = "tb_Mold_DateTime"
        Me.tb_Mold_DateTime.Size = New System.Drawing.Size(68, 21)
        Me.tb_Mold_DateTime.TabIndex = 40
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(1046, 16)
        Me.Label6.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(133, 21)
        Me.Label6.TabIndex = 35
        Me.Label6.Text = "Mold DateTime"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tb_MDL_LOC_ID
        '
        Me.tb_MDL_LOC_ID.BackColor = System.Drawing.SystemColors.Window
        Me.tb_MDL_LOC_ID.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.tb_MDL_LOC_ID.Location = New System.Drawing.Point(978, 85)
        Me.tb_MDL_LOC_ID.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.tb_MDL_LOC_ID.Name = "tb_MDL_LOC_ID"
        Me.tb_MDL_LOC_ID.Size = New System.Drawing.Size(68, 21)
        Me.tb_MDL_LOC_ID.TabIndex = 39
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(845, 85)
        Me.Label5.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(133, 21)
        Me.Label5.TabIndex = 33
        Me.Label5.Text = "MDL LOC ID"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(242, 62)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0, 1, 0, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(133, 21)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Loss Qty"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btn_Search, Me.ToolStripSeparator1, Me.btn_Data_Output, Me.Form_CLose, Me.btn_Table_Show, Me.ToolStripButton1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1264, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btn_Search
        '
        Me.btn_Search.Image = Global.Repair_System.My.Resources.Resources.Search_121
        Me.btn_Search.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Search.Name = "btn_Search"
        Me.btn_Search.Size = New System.Drawing.Size(51, 22)
        Me.btn_Search.Text = "조회"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'btn_Data_Output
        '
        Me.btn_Data_Output.Image = Global.Repair_System.My.Resources.Resources.TEST_FOLDER
        Me.btn_Data_Output.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Data_Output.Name = "btn_Data_Output"
        Me.btn_Data_Output.Size = New System.Drawing.Size(159, 22)
        Me.btn_Data_Output.Text = "반제품 출고 데이터 작성"
        '
        'Form_CLose
        '
        Me.Form_CLose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.Form_CLose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.Form_CLose.Image = Global.Repair_System.My.Resources.Resources.Close
        Me.Form_CLose.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Form_CLose.Name = "Form_CLose"
        Me.Form_CLose.Size = New System.Drawing.Size(23, 22)
        Me.Form_CLose.Text = "폼 닫기"
        '
        'btn_Table_Show
        '
        Me.btn_Table_Show.Image = Global.Repair_System.My.Resources.Resources.folder_black
        Me.btn_Table_Show.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btn_Table_Show.Name = "btn_Table_Show"
        Me.btn_Table_Show.Size = New System.Drawing.Size(83, 22)
        Me.btn_Table_Show.Text = "Table 보기"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GRID_Data_List)
        Me.SplitContainer1.Panel1Collapsed = True
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer1.SplitterDistance = 1229
        Me.SplitContainer1.TabIndex = 10
        '
        'GRID_Data_List
        '
        Me.GRID_Data_List.ColumnInfo = "2,1,0,0,0,100,Columns:"
        Me.GRID_Data_List.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GRID_Data_List.Font = New System.Drawing.Font("굴림", 9.0!)
        Me.GRID_Data_List.Location = New System.Drawing.Point(0, 0)
        Me.GRID_Data_List.Name = "GRID_Data_List"
        Me.GRID_Data_List.Rows.Count = 2
        Me.GRID_Data_List.Rows.DefaultSize = 20
        Me.GRID_Data_List.Size = New System.Drawing.Size(1229, 100)
        Me.GRID_Data_List.StyleInfo = resources.GetString("GRID_Data_List.StyleInfo")
        Me.GRID_Data_List.TabIndex = 55
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.GRID_Lot_List)
        Me.SplitContainer2.Size = New System.Drawing.Size(1264, 748)
        Me.SplitContainer2.SplitterDistance = 127
        Me.SplitContainer2.TabIndex = 0
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(73, 22)
        Me.ToolStripButton1.Text = "'-1' 변환"
        '
        'frm_OMS_Ship_Data
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 773)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Name = "frm_OMS_Ship_Data"
        Me.Text = "OMS 반제품 출고 데이터 생성"
        CType(Me.GRID_Lot_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.GRID_Data_List, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tb_Area_Flag As TextBox
    Friend WithEvents Label21 As Label
    Friend WithEvents tb_To_Site As TextBox
    Friend WithEvents Label23 As Label
    Friend WithEvents tb_Assembly_Site As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents tb_TPN_No As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents tb_Match_Field As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents tb_To_Area As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents tb_From_Area As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents tb_Check_Option As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents tb_Customer_Name As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents tb_Process_ID As TextBox
    Friend WithEvents tb_PCB_Name As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents tb_Start_Tag As TextBox
    Friend WithEvents tb_Work_Week As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents tb_Company_Code As TextBox
    Friend WithEvents tb_Loss_Qty As TextBox
    Friend WithEvents GRID_Lot_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label3 As Label
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents btn_Search As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btn_Data_Output As ToolStripButton
    Friend WithEvents Form_CLose As ToolStripButton
    Friend WithEvents btn_Table_Show As ToolStripButton
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents GRID_Data_List As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents tb_Inkless As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents tb_Mold_DateTime As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents tb_MDL_LOC_ID As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Label10 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents ToolStripButton1 As ToolStripButton
End Class
