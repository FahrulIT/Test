Imports System.Data.OleDb

Public Class frm_monthly_report_of_rw_stock_print

    Dim c_rw As New cls_rw

    'disable exit form
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            Const CS_DBLCLKS As Int32 = &H8
            Const CS_NOCLOSE As Int32 = &H200
            cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
            Return cp
        End Get
    End Property

    Private Sub frm_monthly_report_of_rw_stock_print_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Dispose()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Koneksi()

        '------ngambil tanggal terakhir yang ada datanya
        '------windows 10 (bulan 05/2018 --> problem)

        Dim tglAkhir As String = Nothing
        Dim daD As OleDbDataAdapter
        Dim dsD As New DataSet
        Dim dtD As New DataTable

        Try
            Dim vbl As Integer = CInt(Format(dtDate.Value, "MM-yyyy").Substring(0, 2))
            Dim vth As Integer = CInt(Format(dtDate.Value, "MM-yyyy").Substring(3, 4))
            Dim vtg As Integer = c_rw.getLastDateOfMonth(vbl, vth)
            Dim vtglakhirbulan As String = vtg.ToString & "-" & vbl.ToString & "-" & vth.ToString
            Dim vtglawalbulan As String = "01-" & vbl.ToString & "-" & vth.ToString

            daD = New OleDbDataAdapter("select max(wh_date) as tglAkhir from rw_warehouse " & vbCrLf & _
                  "where wh_date >= to_date('" & vtglawalbulan & "','DD-MM-YYYY') and " & vbCrLf & _
                  "wh_date <= to_date('" & vtglakhirbulan & "','DD-MM-YYYY')", kon)

            daD.Fill(dsD, "tgl_akhir")
            dtD = dsD.Tables("tgl_akhir")

            If dtD.Rows(0)("tglAkhir").ToString = "" Then

            Else
                Dim cmd As OleDbCommand
                cmd = New OleDbCommand("select max(wh_date) as tglAkhir from rw_warehouse " & vbCrLf & _
                      "where wh_date >= to_date('" & vtglawalbulan & "','DD-MM-YYYY') and " & vbCrLf & _
                      "wh_date <= to_date('" & vtglakhirbulan & "','DD-MM-YYYY')", kon)

                Dim tgl_wh As DateTime = DateTime.Now
                tgl_wh = cmd.ExecuteScalar

                cmd = New OleDbCommand("select max(deli_date)as tests from rw_delivery " & vbCrLf & _
                      "where deli_date >= to_date('" & vtglawalbulan & "','DD-MM-YYYY') and " & vbCrLf & _
                      "deli_date <= to_date('" & vtglakhirbulan & "','DD-MM-YYYY')", kon)

                Dim tgl_deli As DateTime = DateTime.Now

                Dim hasil As Object = cmd.ExecuteScalar
                If IsDBNull(hasil) = False Then
                    tgl_deli = hasil
                End If

                If tgl_wh >= tgl_deli Then
                    tglAkhir = tgl_wh.ToString("dd-MM-yyyy")
                Else
                    tglAkhir = tgl_deli.ToString("dd-MM-yyyy")
                End If

            End If

            Dim tglAwal As String = Format(DateAdd(DateInterval.Month, 1, CDate("01/" & _
            Format(DateAdd(DateInterval.Month, -1, dtDate.Value), "MM/yyyy"))), "dd-MM-yyyy")

            Dim dt_add As New DataTable
            dt_add.Columns.Add("cur_date", Type.GetType("System.DateTime"))
            dt_add.Columns.Add("category")
            dt_add.Columns.Add("item_code")
            dt_add.Columns.Add("blended_ratio")
            dt_add.Columns.Add("purc_yarn_name")
            dt_add.Columns.Add("supp_name")
            dt_add.Columns.Add("smm")
            dt_add.Columns.Add("dmm")
            dt_add.Columns.Add("co_hk_pc")
            dt_add.Columns.Add("co_hk_pc_weight", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("dy_fg")
            dt_add.Columns.Add("wh_receive", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("wh_cancel", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("wh_total", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("return_from_dyeing", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("deli_dy", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("deli_fg", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("deli_sp", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("deli_sa", Type.GetType("System.Decimal"))  '[2022.jan.11] spinning sample
            dt_add.Columns.Add("deli_total", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("loss", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("remark")
            dt_add.Columns.Add("balance_last_day", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("balance_today", Type.GetType("System.Decimal"))
            dt_add.Columns.Add("hk_dmtr", Type.GetType("System.Decimal"))
            dt_add.TableName = "dt_daily"

            Dim cur_date As Date = Nothing
            Dim category As String = Nothing
            Dim item_code As String = Nothing
            Dim blended_ratio As String = Nothing
            Dim purc_yarn_name As String = Nothing
            Dim supp_name As String = Nothing
            Dim smm As String = Nothing
            Dim dmm As String = Nothing
            Dim co_hk_pc As String = Nothing
            Dim co_hk_pc_weight As String = Nothing
            Dim dy_fg As String = Nothing
            Dim wh_receive As Decimal = 0
            Dim wh_cancel As Decimal = 0
            Dim wh_total As Decimal = 0
            Dim return_from_dyeing As Decimal = 0
            Dim deli_dy As Decimal = 0
            Dim deli_fg As Decimal = 0
            Dim deli_sp As Decimal = 0
            Dim deli_sa As Decimal = 0
            Dim deli_total As Decimal = 0
            Dim loss As Decimal = 0
            Dim remark As String = Nothing
            Dim balance_last_day As Decimal = 0
            Dim balance_today As Decimal = 0
            Dim hk_dmtr As Decimal = 0
            Dim adaData As Boolean = False

            Dim output As Boolean = False
            If rdScreen.Checked Then
                output = True
            Else
                output = False
            End If

            Dim da As OleDbDataAdapter
            Dim dt_union, dt_stock, dt_last_day As DataTable
            Dim ds As New DataSet
            Dim ds_last_day As New DataSet
            Dim ds_new As New DataSet

            'data akhir bulan
            da = New OleDbDataAdapter("select category, cur_date, rw_type, item_code, blended_ratio, smm, dmm, " & vbCrLf & _
                  "purc_yarn_name, co_hk_pc, co_hk_pc_weight, " & vbCrLf & _
                  "dy_fg, sum(deli_dy) as deli_dy, sum(deli_fg) as deli_fg, sum(deli_sp) as deli_sp, " & vbCrLf & _
                  "sum(deli_sa) as deli_sa,sum(deli_total) as deli_total, remark, sum(wh_receive) as wh_receive," & _
                  "sum(wh_cancel) as wh_cancel, sum(wh_total) as wh_total, " & _
                  "sum(return_from_dyeing) as return_from_dyeing, sum(nvl(loss,0)) as loss, " & vbCrLf & _
                  "supp_name , hk_dmtr from " & vbCrLf & _
                  "( " & vbCrLf & _
                  "select " & vbCrLf & _
                  "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & vbCrLf & _
                  "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & vbCrLf & _
                  "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & vbCrLf & _
                  "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & vbCrLf & _
                  "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & vbCrLf & _
                  "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                  "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                "END END END END END END END END as category, " & vbCrLf & _
                  "h.wh_date as cur_date , h.rw_type, " & vbCrLf & _
                  "case when h.rw_type = 'G' then d.item_name||' SIZE '||d.item_size " & vbCrLf & _
                  "else rtrim(ltrim(h.acm_yarn_kind)) end as item_code, " & vbCrLf & _
                  "case when h.rw_type = 'G' then '' else h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps " & vbCrLf & _
                  "end as blended_ratio, " & vbCrLf & _
                  "h.smm_yarn_count as smm, y.dmm_yarn_count as dmm, rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, " & vbCrLf & _
                  "case when h.rw_type = 'G' then 'PC' else h.wind_type " & vbCrLf & _
                  "end as co_hk_pc, " & vbCrLf & _
                  "case when h.rw_type = 'G' then d.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & vbCrLf & _
                  "h.st_status as dy_fg, 0 as deli_dy, 0 as deli_fg, 0 as deli_sp,0 as deli_sa, 0 as deli_total, " & vbCrLf & _
                  "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, " & vbCrLf & _
                  "case when h.rw_type = 'G' then sum(nvl(d.quantity,0)) else sum(nvl(h.quantity,0)) end as wh_receive, " & vbCrLf & _
                  "0 as wh_cancel, 0 as wh_total, 0 as return_from_dyeing, sum(h.loss) as loss, s.supp_name, y.hk_dmtr " & vbCrLf & _
                  "from rw_warehouse h, rw_warehouse_detail d, yarn_master y, ms_supplier s " & vbCrLf & _
                  "where " & vbCrLf & _
                  "h.wh_slip_no1 = d.wh_slip_no1(+) and " & vbCrLf & _
                  "h.wh_slip_no2 = d.wh_slip_no2(+) and " & vbCrLf & _
                  "h.wh_slip_no3 = d.wh_slip_no3(+) and " & vbCrLf & _
                  "h.wh_slip_no4 = d.wh_slip_no4(+) and " & vbCrLf & _
                  "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & vbCrLf & _
                  "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & vbCrLf & _
                  "rtrim(h.wool_comps) = y.wool_comps(+) and " & vbCrLf & _
                  "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & vbCrLf & _
                  "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & _
                  "h.supp_code = s.supp_code(+) and h.wh_date = to_date('" & tglAkhir & "','DD-MM-YYYY') " & vbCrLf & _
                  "group by h.wh_date, h.rw_type, h.wh_status, d.item_name, d.item_size, rtrim(ltrim(h.acm_yarn_kind)), " & vbCrLf & _
                  "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, " & vbCrLf & _
                  "rtrim(ltrim(h.purc_yarn_name)), h.wind_type, d.pcs_weight, h.co_hk_weight, h.st_status, " & vbCrLf & _
                  "h.RW_STATUS, y.YARN_CLASS, s.supp_name, y.hk_dmtr " & vbCrLf & _
                     "UNION " & vbCrLf & _
                     "select " & vbCrLf & _
                     "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & vbCrLf & _
                     "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                    "END END END END END END END END as category, " & vbCrLf & _
                     "h.cn_date as cur_date, h.rw_type, " & vbCrLf & _
                     "case when h.rw_type = 'G' then d.item_name||' SIZE '||d.item_size else rtrim(ltrim(h.acm_yarn_kind)) " & vbCrLf & _
                     "end as item_code, " & vbCrLf & _
                     "case when h.rw_type = 'G' then '' else h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps " & vbCrLf & _
                     "end as blended_ratio, h.smm_yarn_count as smm, y.dmm_yarn_count as dmm, " & vbCrLf & _
                     "rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, case when h.rw_type = 'G' then 'PC' else h.wind_type " & vbCrLf & _
                     "end as co_hk_pc, " & vbCrLf & _
                     "case when h.rw_type = 'G' then d.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & vbCrLf & _
                     "h.st_status as dy_fg, 0 as deli_dy, 0 as deli_fg, 0 as deli_sp,0 as deli_sa, 0 as deli_total, " & vbCrLf & _
                     "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, 0 as wh_receive, " & vbCrLf & _
                     "case when h.rw_type = 'G' then sum(nvl(d.quantity,0)) else sum(nvl(h.quantity,0)) end as wh_cancel, " & vbCrLf & _
                     "0 as wh_total, 0 as return_from_dyeing, sum(nvl(h.loss,0)) as loss, s.supp_name, y.hk_dmtr " & vbCrLf & _
                     "from rw_cancel_wh h, rw_cancel_wh_detail d, yarn_master y, ms_supplier s " & vbCrLf & _
                     "where " & vbCrLf & _
                     "h.cn_slip_no1 = d.cn_slip_no1(+) and " & vbCrLf & _
                     "h.cn_slip_no2 = d.cn_slip_no2(+) and " & vbCrLf & _
                     "h.cn_slip_no3 = d.cn_slip_no3(+) and " & vbCrLf & _
                     "h.cn_slip_no4 = d.cn_slip_no4(+) and " & vbCrLf & _
                     "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & vbCrLf & _
                     "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & vbCrLf & _
                     "rtrim(h.wool_comps) = y.wool_comps(+) and " & vbCrLf & _
                     "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & vbCrLf & _
                     "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & vbCrLf & _
                     "h.supp_code = s.supp_code(+) and h.cn_date = to_date('" & tglAkhir & "','DD-MM-YYYY') " & vbCrLf & _
                     "group by h.cn_date, h.rw_type, h.wh_status, d.item_name, d.item_size, rtrim(ltrim(h.acm_yarn_kind)), " & vbCrLf & _
                     "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, " & vbCrLf & _
                     "rtrim(ltrim(h.purc_yarn_name)), h.wind_type, d.pcs_weight, h.co_hk_weight, h.st_status, h.RW_STATUS, " & vbCrLf & _
                     "y.YARN_CLASS, s.supp_name, y.hk_dmtr " & vbCrLf & _
                     "UNION " & vbCrLf & _
                     "select " & vbCrLf & _
                     "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & vbCrLf & _
                     "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                     "END END END END END END END END as category, " & vbCrLf & _
                     "h.deli_date as cur_date , h.rw_type, " & vbCrLf & _
                     "case when h.rw_type = 'G' then h.item_name||' SIZE '||h.item_size else rtrim(ltrim(h.acm_yarn_kind)) " & vbCrLf & _
                     "end as item_code, " & vbCrLf & _
                     "case when h.rw_type = 'G' then '' else h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps " & vbCrLf & _
                     "end as blended_ratio, h.smm_yarn_count as smm, y.dmm_yarn_count as dmm, " & vbCrLf & _
                     "rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, " & vbCrLf & _
                     "case when h.rw_type = 'G' then 'PC' else h.wind_type end as co_hk_pc, " & vbCrLf & _
                     "case when h.rw_type = 'G' then h.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & vbCrLf & _
                     "h.st_status as dy_fg, case when h.deli_class = 'DY' then sum(d.quantity) else 0 end as deli_dy, " & vbCrLf & _
                     "case when h.deli_class = 'FG' then sum(d.quantity) else 0 end as deli_fg, " & vbCrLf & _
                     "case when h.deli_class = 'SP' then sum(nvl(d.quantity,0)) else 0 end as deli_sp, " & _
                     "case when h.deli_class = 'SA' then sum(nvl(d.quantity,0)) else 0 end as deli_sa, " & _
                     "0 as deli_total, " & vbCrLf & _
                     "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, 0 as wh_receive, " & vbCrLf & _
                     "0 as wh_cancel, 0 as wh_total, 0 as return_from_dyeing, 0 as loss, s.supp_name, y.hk_dmtr " & vbCrLf & _
                     "from rw_delivery h, rw_delivery_detail d, yarn_master y, ms_supplier s " & vbCrLf & _
                     "where " & vbCrLf & _
                     "h.deli_slip_no1 = d.deli_slip_no1(+) and " & vbCrLf & _
                     "h.deli_slip_no2 = d.deli_slip_no2(+) and " & vbCrLf & _
                     "h.deli_slip_no3 = d.deli_slip_no3(+) and " & vbCrLf & _
                     "h.deli_slip_no4 = d.deli_slip_no4(+) and " & vbCrLf & _
                     "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & vbCrLf & _
                     "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & vbCrLf & _
                     "rtrim(h.wool_comps) = y.wool_comps(+) and " & vbCrLf & _
                     "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & vbCrLf & _
                     "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & vbCrLf & _
                     "h.supp_code = s.supp_code(+) and h.deli_date = to_date('" & tglAkhir & "','DD-MM-YYYY') " & vbCrLf & _
                     "group by h.deli_date, h.rw_type, h.wh_status, h.item_name, h.item_size, rtrim(ltrim(h.acm_yarn_kind))," & vbCrLf & _
                     "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count," & vbCrLf & _
                     "rtrim(ltrim(h.purc_yarn_name)), h.wind_type, h.pcs_weight, h.co_hk_weight, h.st_status, h.deli_class," & vbCrLf & _
                     "h.RW_STATUS, y.YARN_CLASS, s.supp_name, y.hk_dmtr " & _
                     "UNION " & vbCrLf & _
                     "select " & vbCrLf & _
                     "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & vbCrLf & _
                     "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                    "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                    "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                    "END END END END END END END END as category, " & vbCrLf & _
                     "h.wh_date as cur_date , h.rw_type, rtrim(ltrim(h.acm_yarn_kind)) as item_code," & vbCrLf & _
                     "h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps as blended_ratio, h.smm_yarn_count as smm," & vbCrLf & _
                     "y.dmm_yarn_count as dmm, rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, h.wind_type as co_hk_pc," & vbCrLf & _
                     "h.co_hk_weight as co_hk_pc_weight, h.st_status as dy_fg, 0 as deli_dy, 0 as deli_fg, 0 as deli_sp," & vbCrLf & _
                     "0 as deli_sa,0 as deli_total, " & vbCrLf & _
                     "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, 0 as wh_receive, " & vbCrLf & _
                     "0 as wh_cancel, 0 as wh_total, sum(nvl(h.quantity,0))as return_from_dyeing, 0 as loss, s.supp_name," & vbCrLf & _
                     "y.hk_dmtr " & vbCrLf & _
                     "from rw_return h, yarn_master y, ms_supplier s " & vbCrLf & _
                     "where " & vbCrLf & _
                     "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & vbCrLf & _
                     "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & vbCrLf & _
                     "rtrim(h.wool_comps) = y.wool_comps(+) and " & vbCrLf & _
                     "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & vbCrLf & _
                     "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & vbCrLf & _
                     "h.supp_code = s.supp_code(+) and h.wh_date = to_date('" & tglAkhir & "','DD-MM-YYYY') " & vbCrLf & _
                     "group by h.wh_date, h.rw_type, h.wh_status, rtrim(ltrim(h.acm_yarn_kind)), h.acrylic_comps, " & vbCrLf & _
                     "h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, rtrim(ltrim(h.purc_yarn_name))," & vbCrLf & _
                     "h.wind_type, h.co_hk_weight, h.st_status, h.RW_STATUS, y.YARN_CLASS, s.supp_name, y.hk_dmtr " & vbCrLf & _
                     "UNION ALL " & vbCrLf & _
                     "select " & vbCrLf & _
                     "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & vbCrLf & _
                     "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                     "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                    "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                    "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                    "END END END END END END END END as category, " & vbCrLf & _
                     "h.wh_date as cur_date, h.rw_type, " & vbCrLf & _
                     "case when h.rw_type = 'G' then d.item_name||' SIZE '||d.item_size else rtrim(ltrim(h.acm_yarn_kind)) " & vbCrLf & _
                     "end as item_code, " & vbCrLf & _
                     "case when h.rw_type = 'G' then '' else h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps " & vbCrLf & _
                     "end as blended_ratio, h.smm_yarn_count as smm, y.dmm_yarn_count as dmm," & vbCrLf & _
                     "rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, " & vbCrLf & _
                     "case when h.rw_type = 'G' then 'PC' else h.wind_type end as co_hk_pc, " & vbCrLf & _
                     "case when h.rw_type = 'G' then d.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & vbCrLf & _
                     "h.st_status as dy_fg,0 as deli_dy,0 as deli_fg,0 as deli_sp,0 as deli_sa,0 as deli_total, " & vbCrLf & _
                     "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, " & vbCrLf & _
                     "case when h.rw_type = 'G' then sum(nvl(d.quantity,0)) else sum(nvl(h.quantity,0)) end as wh_receive," & vbCrLf & _
                     "0 as wh_cancel, 0 as wh_total, 0 as return_from_dyeing, sum(nvl(h.loss,0)) as loss, s.supp_name, y.hk_dmtr " & vbCrLf & _
                     "from rw_cancel_wh h, rw_cancel_wh_detail d, yarn_master y, ms_supplier s " & vbCrLf & _
                     "where " & vbCrLf & _
                     "h.cn_slip_no1 = d.cn_slip_no1(+) and " & vbCrLf & _
                     "h.cn_slip_no2 = d.cn_slip_no2(+) and " & vbCrLf & _
                     "h.cn_slip_no3 = d.cn_slip_no3(+) and " & vbCrLf & _
                     "h.cn_slip_no4 = d.cn_slip_no4(+) and " & vbCrLf & _
                     "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & vbCrLf & _
                     "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & vbCrLf & _
                     "rtrim(h.wool_comps) = y.wool_comps(+) and " & vbCrLf & _
                     "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & vbCrLf & _
                     "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & vbCrLf & _
                     "h.supp_code = s.supp_code(+) and h.wh_date = to_date('" & tglAkhir & "','DD-MM-YYYY') " & vbCrLf & _
                     "group by h.wh_date, h.rw_type, h.wh_status, d.item_name, d.item_size, rtrim(ltrim(h.acm_yarn_kind))," & vbCrLf & _
                     "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count," & vbCrLf & _
                     "rtrim(ltrim(h.purc_yarn_name)), h.wind_type, d.pcs_weight, h.co_hk_weight, h.st_status, h.RW_STATUS," & vbCrLf & _
                     "y.YARN_CLASS, s.supp_name, y.hk_dmtr) " & vbCrLf & _
                     "group by category, cur_date, rw_type, item_code, blended_ratio, smm, dmm, purc_yarn_name, co_hk_pc," & vbCrLf & _
                     "co_hk_pc_weight, dy_fg, remark,  supp_name, hk_dmtr " & vbCrLf & _
                     "order by item_code", kon)

            da.Fill(ds, "dt_all")
            dt_union = ds.Tables("dt_all") 'data akhir bulan

            For i As Integer = 0 To dt_union.Rows.Count - 1
                'If dt_union.Rows(i)("item_code").ToString = "POLY150D/48F" And dt_union.Rows(i)("smm").ToString() = "2/60" Then
                '    Dim a As String = ""
                'End If
                cur_date = dt_union.Rows(i)("cur_date")
                category = dt_union.Rows(i)("category").ToString()
                item_code = dt_union.Rows(i)("item_code").ToString()
                blended_ratio = dt_union.Rows(i)("blended_ratio").ToString()
                purc_yarn_name = dt_union.Rows(i)("purc_yarn_name").ToString()
                supp_name = dt_union.Rows(i)("supp_name").ToString()
                smm = dt_union.Rows(i)("smm").ToString()
                dmm = dt_union.Rows(i)("dmm").ToString()
                co_hk_pc = dt_union.Rows(i)("co_hk_pc").ToString()
                co_hk_pc_weight = IIf(dt_union.Rows(i)("co_hk_pc_weight").ToString() = "", 0, dt_union.Rows(i)("co_hk_pc_weight").ToString())
                dy_fg = dt_union.Rows(i)("dy_fg").ToString()
                remark = dt_union.Rows(i)("remark").ToString()
                wh_receive = 0
                wh_cancel = 0
                wh_total = 0
                return_from_dyeing = 0
                deli_dy = 0
                deli_fg = 0
                deli_sp = 0
                deli_sa = 0 '[2022.jan.11] spinning-sample
                deli_total = 0
                loss = dt_union.Rows(i)("loss").ToString() '0
                hk_dmtr = IIf(dt_union.Rows(i)("hk_dmtr").ToString() = "", 0, dt_union.Rows(i)("hk_dmtr").ToString())
                adaData = False

                For Each row As DataRow In dt_union.Rows
                    If category = row("category").ToString And item_code = row("item_code").ToString And _
                    blended_ratio = row("blended_ratio").ToString And purc_yarn_name = row("purc_yarn_name").ToString And _
                    supp_name = row("supp_name").ToString And smm = row("smm").ToString And dmm = row("dmm").ToString And _
                    co_hk_pc = row("co_hk_pc").ToString And co_hk_pc_weight = row("co_hk_pc_weight").ToString And _
                    dy_fg = row("dy_fg").ToString And remark = row("remark").ToString Then
                        wh_receive += row("wh_receive").ToString
                        wh_cancel += row("wh_cancel").ToString
                        wh_total = wh_receive - wh_cancel
                        return_from_dyeing += row("return_from_dyeing").ToString
                        deli_dy += row("deli_dy").ToString
                        deli_fg += row("deli_fg").ToString
                        deli_sp += row("deli_sp").ToString
                        deli_sa += row("deli_sa").ToString                  '[2022.jan.11] spinning-sample
                        deli_total = deli_dy + deli_fg + deli_sp + deli_sa  '[2022.jan.11] spinning-sample
                        'loss += row("loss").ToString
                    End If
                Next

                For Each rowSame As DataRow In dt_add.Rows
                    'If rowSame("item_code").ToString = "POLY150D/48F" And rowSame("smm").ToString() = "2/60" Then
                    '    Dim a As String = ""
                    'End If
                    If category = rowSame("category").ToString And item_code = rowSame("item_code").ToString And _
                    blended_ratio = rowSame("blended_ratio").ToString And purc_yarn_name = rowSame("purc_yarn_name").ToString And _
                    supp_name = rowSame("supp_name").ToString And smm = rowSame("smm").ToString And _
                    dmm = rowSame("dmm").ToString And co_hk_pc = rowSame("co_hk_pc").ToString And _
                    co_hk_pc_weight = rowSame("co_hk_pc_weight").ToString And dy_fg = rowSame("dy_fg").ToString And _
                    remark = rowSame("remark").ToString Then
                        adaData = True
                    End If
                Next

                'add to dt_add
                If adaData = False Then
                    Dim addRow As DataRow = dt_add.NewRow
                    addRow("cur_date") = cur_date
                    addRow("category") = category
                    addRow("item_code") = item_code
                    addRow("blended_ratio") = blended_ratio
                    addRow("purc_yarn_name") = purc_yarn_name
                    addRow("supp_name") = supp_name
                    addRow("smm") = smm
                    addRow("dmm") = dmm
                    addRow("co_hk_pc") = co_hk_pc
                    addRow("co_hk_pc_weight") = co_hk_pc_weight
                    addRow("dy_fg") = dy_fg
                    addRow("wh_receive") = wh_receive
                    addRow("wh_cancel") = wh_cancel
                    addRow("wh_total") = wh_total
                    addRow("return_from_dyeing") = return_from_dyeing
                    addRow("deli_dy") = deli_dy
                    addRow("deli_fg") = deli_fg
                    addRow("deli_sp") = deli_sp
                    addRow("deli_sa") = deli_sa '[2022.jan.11] spinning-sample
                    addRow("deli_total") = deli_total
                    addRow("loss") = loss
                    addRow("remark") = remark
                    addRow("hk_dmtr") = hk_dmtr
                    dt_add.Rows.Add(addRow)
                End If
            Next

            ds.Tables.Remove("dt_all") 'remove dt_all
            ds.Tables.Add(dt_add)      ' add dt_daily

            'cek delivery
            'Dim newDataTable As DataTable = dt_add.Clone
            'Dim dataRows As DataRow() = dt_add.Select("item_code='POLY150D/48F' and smm='2/60'", "")
            'Dim dr As DataRow
            'For Each dr In dataRows
            '    newDataTable.ImportRow(dr)
            'Next

            Dim blth As String = Format(dtDate.Value, "MM-yyyy")

            Dim n_bl As Integer = CInt(blth.Substring(0, 2))
            Dim n_th As Integer = CInt(blth.Substring(3, 4))
            Dim bl, th As String
            If n_bl = 1 Then
                bl = "12"
                th = (n_th - 1).ToString
            Else
                bl = (n_bl - 1).ToString.PadLeft(2, "0")
                th = blth.Substring(3, 4)
            End If

            '-- RW Stock
            da = New OleDbDataAdapter("SELECT CATEGORY, month, year, rw_type, item_code, blended_ratio, smm, dmm," & _
                   "purc_yarn_name, co_hk_pc, co_hk_pc_weight, " & _
                   "dy_fg, remark, supp_name, sum(BALANCE_LAST_DAY)as BALANCE_LAST_DAY, hk_dmtr from (select " & _
                   "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & _
                   "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & _
                   "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & _
                   "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & _
                   "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & _
                   "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                   "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                    "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                    "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                    "END END END END END END END END as category, " & vbCrLf & _
                   "h.month, h.year, h.rw_type, case when h.rw_type = 'G' then h.item_name||' SIZE '||" & _
                   "h.item_size else rtrim(ltrim(h.acm_yarn_kind)) end as item_code, " & _
                   "case when h.rw_type = 'G' then '' else h.acrylic_comps||'/'||h.wool_comps||'/'||" & _
                   "h.nylon_comps end as blended_ratio, h.smm_yarn_count as smm, y.dmm_yarn_count as dmm, " & _
                   "h.purc_yarn_name, h.wind_type as co_hk_pc, case when h.rw_type = 'G' then h.pcs_weight else " & _
                   "h.co_hk_weight end as co_hk_pc_weight, h.st_status as dy_fg, " & _
                   "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, s.supp_name, " & _
                   "sum(nvl(h.quantity,0) + nvl(h.quantity_a,0))as balance_last_day, y.hk_dmtr " & _
                   "from rw_stock h, yarn_master y, ms_supplier s " & _
                   "where " & _
                   "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & _
                   "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & _
                   "rtrim(h.wool_comps) = y.wool_comps(+) and " & _
                   "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & _
                   "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & _
                   "h.supp_code = s.supp_code(+) and " & _
                   "h.month = '" & bl & "' and  " & _
                   "h.year = '" & th & "' " & _
                   "group by h.month, h.year, h.rw_type, h.wh_status, rtrim(ltrim(h.acm_yarn_kind)), h.acrylic_comps," & _
                   "h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, h.purc_yarn_name, h.wind_type," & _
                   "h.co_hk_weight, h.st_status, h.RW_STATUS, y.YARN_CLASS, s.supp_name, h.item_name, h.item_size," & _
                   "h.pcs_weight, y.hk_dmtr " & _
                   ") " & _
                   "group by CATEGORY, month, year, rw_type, item_code, blended_ratio, smm, dmm, purc_yarn_name," & _
                   "co_hk_pc, co_hk_pc_weight,dy_fg, remark, supp_name, hk_dmtr ", kon)

            da.Fill(ds, "dt_stock")
            dt_stock = ds.Tables("dt_stock")


            If dtD.Rows(0)("tglAkhir").ToString = "" Then
                '----tambah stock ke dt_add
                For a As Integer = 0 To dt_stock.Rows.Count - 1
                    cur_date = dtDate.Value
                    category = dt_stock.Rows(a)("category").ToString()
                    item_code = dt_stock.Rows(a)("item_code").ToString()
                    blended_ratio = dt_stock.Rows(a)("blended_ratio").ToString()
                    purc_yarn_name = dt_stock.Rows(a)("purc_yarn_name").ToString()
                    supp_name = dt_stock.Rows(a)("supp_name").ToString()
                    smm = dt_stock.Rows(a)("smm").ToString()
                    dmm = dt_stock.Rows(a)("dmm").ToString()
                    co_hk_pc = dt_stock.Rows(a)("co_hk_pc").ToString()
                    co_hk_pc_weight = dt_stock.Rows(a)("co_hk_pc_weight").ToString()
                    dy_fg = dt_stock.Rows(a)("dy_fg").ToString()
                    remark = dt_stock.Rows(a)("remark").ToString()
                    'loss = dt_stock.Rows(a)("loss").ToString()
                    balance_last_day = dt_stock.Rows(a)("balance_last_day").ToString()
                    hk_dmtr = dt_stock.Rows(a)("hk_dmtr").ToString()
                    adaData = False

                    For Each row As DataRow In dt_add.Rows
                        If category = row("category").ToString And item_code = row("item_code").ToString And _
                           blended_ratio = row("blended_ratio").ToString And purc_yarn_name = row("purc_yarn_name").ToString And _
                           supp_name = row("supp_name").ToString And smm = row("smm").ToString And _
                           dmm = row("dmm").ToString And co_hk_pc = row("co_hk_pc").ToString And _
                           co_hk_pc_weight = row("co_hk_pc_weight").ToString And dy_fg = row("dy_fg").ToString And _
                           remark = row("remark").ToString Then
                            row("balance_last_day") = balance_last_day
                            row("balance_today") = balance_last_day + row("wh_total") + row("return_from_dyeing") - row("deli_total") '- row("loss")
                            adaData = True
                        End If
                    Next

                    'nambah stock ke dt_add
                    If adaData = False Then
                        Dim stAdd As DataRow = dt_add.NewRow
                        stAdd("cur_date") = cur_date
                        stAdd("category") = category
                        stAdd("item_code") = item_code
                        stAdd("blended_ratio") = blended_ratio
                        stAdd("purc_yarn_name") = purc_yarn_name
                        stAdd("supp_name") = supp_name
                        stAdd("smm") = smm
                        stAdd("dmm") = dmm
                        stAdd("co_hk_pc") = co_hk_pc
                        stAdd("co_hk_pc_weight") = co_hk_pc_weight
                        stAdd("dy_fg") = dy_fg
                        stAdd("wh_receive") = 0
                        stAdd("wh_cancel") = 0
                        stAdd("wh_total") = 0
                        stAdd("return_from_dyeing") = 0
                        stAdd("deli_dy") = 0
                        stAdd("deli_fg") = 0
                        stAdd("deli_sp") = 0
                        stAdd("deli_sa") = 0  '[2022.jan.11] spinning-sample
                        stAdd("deli_total") = 0
                        stAdd("loss") = 0
                        stAdd("remark") = remark
                        stAdd("balance_last_day") = balance_last_day
                        stAdd("balance_today") = balance_last_day
                        stAdd("hk_dmtr") = hk_dmtr
                        dt_add.Rows.Add(stAdd)
                    End If
                Next
                For Each rAdd As DataRow In dt_add.Rows
                    If rAdd("balance_today").ToString = "" Then
                        rAdd("balance_today") = rAdd("wh_total") + rAdd("return_from_dyeing") - rAdd("deli_total") '- rAdd("loss")
                    End If
                Next

            Else
                'ADD DATA TANGGAL KEMAREN 
                Dim tg_k As String = ""
                tg_k = Format(DateAdd(DateInterval.Day, -1, CDate(tglAkhir)), "dd-MM-yyyy")
                'tg_k = "30-10-2019"

                da = New OleDbDataAdapter("select category, rw_type, item_code, blended_ratio, smm, dmm, " & vbCrLf & _
                        "purc_yarn_name, co_hk_pc, co_hk_pc_weight, " & vbCrLf & _
                        "dy_fg, sum(deli_dy) as deli_dy, sum(deli_fg) as deli_fg, sum(deli_sp) as deli_sp, " & vbCrLf & _
                        "sum(deli_sa) as deli_sa, sum(deli_total) as deli_total, remark, sum(wh_receive) as wh_receive, " & vbCrLf & _
                        "sum(wh_cancel) as wh_cancel, sum(wh_total) as wh_total, sum(return_from_dyeing) as return_from_dyeing, " & _
                        "sum(loss) as loss, supp_name, sum(balance_last_day) as balance_last_day, " & _
                        "sum(balance_today) as balance_today, hk_dmtr from ( " & _
                        "select " & _
                        "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & _
                        "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & _
                        "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                        "END END END END END END END END as category, " & vbCrLf & _
                        "h.rw_type, case when h.rw_type = 'G' then d.item_name||' SIZE '||d.item_size else " & _
                        "rtrim(ltrim(h.acm_yarn_kind)) end as item_code, case when h.rw_type = 'G' then '' else " & _
                        "h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps end as blended_ratio, " & _
                        "h.smm_yarn_count as smm, y.dmm_yarn_count as dmm, rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, " & _
                        "case when h.rw_type = 'G' then 'PC' else h.wind_type end as co_hk_pc, " & _
                        "case when h.rw_type = 'G' then d.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & _
                        "h.st_status as dy_fg, 0 as deli_dy, 0 as deli_fg, 0 as deli_sp,0 as deli_sa,0 as deli_total, " & _
                        "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, " & _
                        "case when h.rw_type = 'G' then sum(nvl(d.quantity,0)) else sum(nvl(h.quantity,0)) end as wh_receive," & _
                        "0 as wh_cancel, 0 as wh_total, 0 as return_from_dyeing, sum(nvl(h.loss,0)) as loss, s.supp_name," & _
                        "0 AS balance_last_day, 0 AS balance_today, y.hk_dmtr " & _
                        "from rw_warehouse h, rw_warehouse_detail d, yarn_master y, ms_supplier s " & _
                        "where " & _
                        "h.wh_slip_no1 = d.wh_slip_no1(+) and " & _
                        "h.wh_slip_no2 = d.wh_slip_no2(+) and " & _
                        "h.wh_slip_no3 = d.wh_slip_no3(+) and " & _
                        "h.wh_slip_no4 = d.wh_slip_no4(+) and " & _
                        "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & _
                        "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & _
                        "rtrim(h.wool_comps) = y.wool_comps(+) and " & _
                        "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & _
                        "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & vbCrLf & _
                        "h.supp_code = s.supp_code(+) and " & _
                        "h.wh_date >= to_date('" & tglAwal & "','DD-MM-YYYY') and h.wh_date <= to_date('" & tg_k & "','DD-MM-YYYY') " & vbCrLf & _
                        "group by h.rw_type, h.wh_status, d.item_name, d.item_size, rtrim(ltrim(h.acm_yarn_kind)), " & _
                        "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, " & _
                        "h.purc_yarn_name, h.wind_type, d.pcs_weight, h.co_hk_weight, h.st_status, h.RW_STATUS, " & _
                        "y.YARN_CLASS, s.supp_name, y.hk_dmtr " & _
                        "UNION " & _
                        "select " & _
                        "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & _
                        "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & _
                        "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                        "END END END END END END END END as category, " & vbCrLf & _
                        "h.rw_type, case when h.rw_type = 'G' then d.item_name||' SIZE '||" & _
                        "d.item_size else rtrim(ltrim(h.acm_yarn_kind)) end as item_code, " & _
                        "case when h.rw_type = 'G' then '' else " & _
                        "h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps end as blended_ratio, " & _
                        "h.smm_yarn_count as smm, y.dmm_yarn_count as dmm, rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name," & _
                        "case when h.rw_type = 'G' then 'PC' else h.wind_type end as co_hk_pc, " & _
                        "case when h.rw_type = 'G' then d.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & _
                        "h.st_status as dy_fg, 0 as deli_dy, 0 as deli_fg, 0 as deli_sp,0 as deli_sa,0 as deli_total, " & _
                        "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, 0 as wh_receive, " & _
                        "case when h.rw_type = 'G' then sum(nvl(d.quantity,0)) else sum(nvl(h.quantity,0)) end as wh_cancel," & _
                        "0 as wh_total, 0 as return_from_dyeing, sum(nvl(h.loss,0)) as loss, s.supp_name, " & _
                        "0 AS balance_last_day, 0 AS balance_today, y.hk_dmtr " & _
                        "from rw_cancel_wh h, rw_cancel_wh_detail d, yarn_master y, ms_supplier s " & _
                        "where " & _
                        "h.cn_slip_no1 = d.cn_slip_no1(+) and " & _
                        "h.cn_slip_no2 = d.cn_slip_no2(+) and " & _
                        "h.cn_slip_no3 = d.cn_slip_no3(+) and " & _
                        "h.cn_slip_no4 = d.cn_slip_no4(+) and " & _
                        "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & _
                        "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & _
                        "rtrim(h.wool_comps) = y.wool_comps(+) and " & _
                        "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & _
                        "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & vbCrLf & _
                        "h.supp_code = s.supp_code(+) and " & vbCrLf & _
                        "h.cn_date >= to_date('" & tglAwal & "','DD-MM-YYYY') and h.cn_date <= to_date('" & tg_k & "','DD-MM-YYYY') " & vbCrLf & _
                        "group by h.rw_type, h.wh_status, d.item_name, d.item_size, rtrim(ltrim(h.acm_yarn_kind)), " & _
                        "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, " & _
                        "h.purc_yarn_name, h.wind_type, d.pcs_weight, h.co_hk_weight, h.st_status, h.RW_STATUS, " & _
                        "y.YARN_CLASS, s.supp_name, y.hk_dmtr " & _
                        "UNION " & _
                        "select " & _
                        "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & _
                        "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & _
                        "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                        "END END END END END END END END as category, " & vbCrLf & _
                        "h.rw_type, case when h.rw_type = 'G' then h.item_name||' SIZE '||h.item_size else " & _
                        "rtrim(ltrim(h.acm_yarn_kind)) end as item_code, case when h.rw_type = 'G' then '' else " & _
                        "h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps end as blended_ratio, h.smm_yarn_count as smm, " & _
                        "y.dmm_yarn_count as dmm, rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, " & _
                        "case when h.rw_type = 'G' then 'PC' else h.wind_type end as co_hk_pc, " & _
                        "case when h.rw_type = 'G' then h.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & _
                        "h.st_status as dy_fg, case when h.deli_class = 'DY' then sum(d.quantity) else 0 end as deli_dy, " & _
                        "case when h.deli_class = 'FG' then sum(d.quantity) else 0 end as deli_fg, " & _
                        "case when h.deli_class = 'SP' then sum(nvl(d.quantity,0)) else 0 end as deli_sp, " & _
                        "case when h.deli_class = 'SA' then sum(nvl(d.quantity,0)) else 0 end as deli_sa, " & _
                        "0 as deli_total, " & _
                        "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, 0 as wh_receive, " & _
                        "0 as wh_cancel, 0 as wh_total, 0 as return_from_dyeing, 0 as loss, s.supp_name, 0 AS balance_last_day, " & _
                        "0 AS balance_today, y.hk_dmtr " & _
                        "from rw_delivery h, rw_delivery_detail d, yarn_master y, ms_supplier s " & _
                        "where " & _
                        "h.deli_slip_no1 = d.deli_slip_no1(+) and " & _
                        "h.deli_slip_no2 = d.deli_slip_no2(+) and " & _
                        "h.deli_slip_no3 = d.deli_slip_no3(+) and " & _
                        "h.deli_slip_no4 = d.deli_slip_no4(+) and " & _
                        "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & _
                        "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & _
                        "rtrim(h.wool_comps) = y.wool_comps(+) and " & _
                        "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & _
                        "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & _
                        "h.supp_code = s.supp_code(+) and " & vbCrLf & _
                        "h.deli_date >= to_date('" & tglAwal & "','DD-MM-YYYY') and h.deli_date <= to_date('" & tg_k & "','DD-MM-YYYY') " & vbCrLf & _
                        "group by h.rw_type, h.wh_status, h.item_name, h.item_size, rtrim(ltrim(h.acm_yarn_kind)), " & _
                        "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, " & _
                        "rtrim(ltrim(h.purc_yarn_name)), h.wind_type, h.pcs_weight, h.co_hk_weight, h.st_status, h.deli_class, " & _
                        "h.RW_STATUS, y.YARN_CLASS, s.supp_name, y.hk_dmtr " & _
                        "UNION " & _
                        "select " & _
                        "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & _
                        "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & _
                        "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                        "END END END END END END END END as category, " & vbCrLf & _
                        "h.rw_type, rtrim(ltrim(h.acm_yarn_kind)) as item_code, " & _
                        "h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps as blended_ratio, h.smm_yarn_count as smm, " & _
                        "y.dmm_yarn_count as dmm, rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, h.wind_type as co_hk_pc, " & _
                        "h.co_hk_weight as co_hk_pc_weight, h.st_status as dy_fg, 0 as deli_dy, 0 as deli_fg, 0 as deli_sp, " & _
                        "0 as deli_sa,0 as deli_total, case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, " & _
                        "0 as wh_receive, 0 as wh_cancel, 0 as wh_total, sum(nvl(h.quantity,0))as return_from_dyeing, " & _
                        "0 as loss, s.supp_name, 0 AS balance_last_day, 0 AS balance_today, y.hk_dmtr " & _
                        "from rw_return h, yarn_master y, ms_supplier s " & _
                        "where " & _
                        "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & _
                        "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & _
                        "rtrim(h.wool_comps) = y.wool_comps(+) and " & _
                        "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & _
                        "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & _
                        "h.supp_code = s.supp_code(+) and " & _
                        "h.wh_date >= to_date('" & tglAwal & "','DD-MM-YYYY') and h.wh_date <= to_date('" & tg_k & "','DD-MM-YYYY') " & vbCrLf & _
                        "group by h.rw_type, h.wh_status, rtrim(ltrim(h.acm_yarn_kind)), h.acrylic_comps, h.wool_comps," & _
                        "h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, h.purc_yarn_name, h.wind_type," & _
                        "h.co_hk_weight, h.st_status, h.RW_STATUS, y.YARN_CLASS, s.supp_name, y.hk_dmtr " & _
                        "UNION ALL " & _
                        "select " & _
                        "CASE WHEN h.rw_type = 'Y' AND h.RW_STATUS = 'P' THEN '6. PURCHASE' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'R' THEN '1. REGULAR' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'S' THEN '2. SPECIAL' ELSE " & _
                        "CASE WHEN h.WH_STATUS IN ('1','4') AND y.YARN_CLASS = 'W' THEN '3. WOOL/BLENDED'  ELSE " & _
                        "CASE WHEN h.WH_STATUS = '2' THEN '4. ABNORMAL' ELSE " & _
                        "--CASE WHEN h.WH_STATUS = '3' THEN '5. TEST' else '7. GARMENT' END END END END END END as category, " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'R' THEN '5.a. TEST-REGULAR' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'W' THEN '5.b. TEST-WOOL' ELSE " & vbCrLf & _
                        "CASE WHEN h.WH_STATUS = '3' AND y.YARN_CLASS = 'S' THEN '5.c. TEST-SPECIAL' ELSE '7. GARMENT' " & vbCrLf & _
                        "END END END END END END END END as category, " & vbCrLf & _
                        "h.rw_type, case when h.rw_type = 'G' then d.item_name||' SIZE '||d.item_size else " & _
                        "rtrim(ltrim(h.acm_yarn_kind)) end as item_code, case when h.rw_type = 'G' then '' else " & _
                        "h.acrylic_comps||'/'||h.wool_comps||'/'||h.nylon_comps end as blended_ratio, h.smm_yarn_count as smm, " & _
                        "y.dmm_yarn_count as dmm, rtrim(ltrim(h.purc_yarn_name))as purc_yarn_name, " & _
                        "case when h.rw_type = 'G' then 'PC' else h.wind_type end as co_hk_pc, " & _
                        "case when h.rw_type = 'G' then d.pcs_weight else h.co_hk_weight end as co_hk_pc_weight, " & _
                        "h.st_status as dy_fg, 0 as deli_dy, 0 as deli_fg, 0 as deli_sp,0 as deli_sa, 0 as deli_total, " & _
                        "case when h.wh_status = '4' then 'ABNORMAL WEIGHT' else '' end as remark, " & _
                        "case when h.rw_type = 'G' then sum(nvl(d.quantity,0)) else sum(nvl(h.quantity,0)) end as wh_receive, " & _
                        "0 as wh_cancel, 0 as wh_total, 0 as return_from_dyeing, sum(nvl(h.loss,0)) as loss, s.supp_name," & _
                        "0 AS balance_last_day, 0 AS balance_today, y.hk_dmtr " & _
                        "from rw_cancel_wh h, rw_cancel_wh_detail d, yarn_master y, ms_supplier s " & _
                        "where " & _
                        "h.cn_slip_no1 = d.cn_slip_no1(+) and " & _
                        "h.cn_slip_no2 = d.cn_slip_no2(+) and " & _
                        "h.cn_slip_no3 = d.cn_slip_no3(+) and " & _
                        "h.cn_slip_no4 = d.cn_slip_no4(+) and " & _
                        "rtrim(h.acm_yarn_kind) = y.acm_yarn_kind(+) and " & _
                        "rtrim(h.acrylic_comps) = y.acrylic_comps(+) and " & _
                        "rtrim(h.wool_comps) = y.wool_comps(+) and " & _
                        "rtrim(h.nylon_comps) = y.nylon_comps(+) and " & _
                        "rtrim(h.smm_yarn_count) = y.smm_yarn_count(+) and " & _
                        "h.supp_code = s.supp_code(+) and " & _
                        "h.wh_date >= to_date('" & tglAwal & "','DD-MM-YYYY') and " & _
                        "h.wh_date <= to_date('" & tg_k & "','DD-MM-YYYY') " & vbCrLf & _
                        "group by h.rw_type, h.wh_status, d.item_name, d.item_size, rtrim(ltrim(h.acm_yarn_kind)), " & _
                        "h.acrylic_comps, h.wool_comps, h.nylon_comps, h.smm_yarn_count, y.dmm_yarn_count, " & _
                        "h.purc_yarn_name, h.wind_type, d.pcs_weight, h.co_hk_weight, h.st_status, h.RW_STATUS, " & _
                        "y.YARN_CLASS, s.supp_name, y.hk_dmtr ) " & _
                        "group by category, rw_type, item_code, blended_ratio, smm, dmm, purc_yarn_name, co_hk_pc, co_hk_pc_weight, " & _
                        "dy_fg, remark,  supp_name, hk_dmtr " & _
                        "order by item_code", kon)

                da.Fill(ds_last_day, "dt_last_day ")
                dt_last_day = ds_last_day.Tables("dt_last_day ")

                'cek data
                'Dim newDataTable As DataTable = dt_last_day.Clone
                'Dim dataRows As DataRow() = dt_last_day.Select("item_code like '%POLY150D/48F%' and smm='2/60'", "")
                'Dim dr As DataRow
                'For Each dr In dataRows
                '    newDataTable.ImportRow(dr)
                'Next

                'gabungin dgn stock tanggal kemarin
                For a As Integer = 0 To dt_stock.Rows.Count - 1
                    category = dt_stock.Rows(a)("category").ToString()
                    item_code = dt_stock.Rows(a)("item_code").ToString()
                    blended_ratio = dt_stock.Rows(a)("blended_ratio").ToString()
                    purc_yarn_name = dt_stock.Rows(a)("purc_yarn_name").ToString()
                    supp_name = dt_stock.Rows(a)("supp_name").ToString()
                    smm = dt_stock.Rows(a)("smm").ToString()
                    dmm = dt_stock.Rows(a)("dmm").ToString()
                    co_hk_pc = dt_stock.Rows(a)("co_hk_pc").ToString()
                    co_hk_pc_weight = dt_stock.Rows(a)("co_hk_pc_weight").ToString()
                    dy_fg = dt_stock.Rows(a)("dy_fg").ToString()
                    remark = dt_stock.Rows(a)("remark").ToString()
                    'loss = dt_stock.Rows(a)("loss").ToString()
                    balance_last_day = dt_stock.Rows(a)("balance_last_day").ToString()
                    hk_dmtr = IIf(dt_stock.Rows(a)("hk_dmtr").ToString() = "", 0, dt_stock.Rows(a)("hk_dmtr").ToString())
                    adaData = False

                    For Each row As DataRow In dt_last_day.Rows
                        If category = row("category").ToString And item_code = row("item_code").ToString And blended_ratio = row("blended_ratio").ToString And purc_yarn_name = row("purc_yarn_name").ToString And supp_name = row("supp_name").ToString And smm = row("smm").ToString And dmm = row("dmm").ToString And co_hk_pc = row("co_hk_pc").ToString And co_hk_pc_weight = row("co_hk_pc_weight").ToString And dy_fg = row("dy_fg").ToString And remark = row("remark").ToString Then
                            row("balance_last_day") = balance_last_day
                            row("balance_today") = balance_last_day + row("wh_total") + row("return_from_dyeing") - row("deli_total") '- row("loss")
                            adaData = True
                        End If
                    Next

                    'nambah stock ke dt_add
                    If adaData = False Then
                        Dim stAdd As DataRow = dt_last_day.NewRow
                        stAdd("category") = category
                        stAdd("item_code") = item_code
                        stAdd("blended_ratio") = blended_ratio
                        stAdd("purc_yarn_name") = purc_yarn_name
                        stAdd("supp_name") = supp_name
                        stAdd("smm") = smm
                        stAdd("dmm") = dmm
                        stAdd("co_hk_pc") = co_hk_pc
                        stAdd("co_hk_pc_weight") = co_hk_pc_weight
                        stAdd("dy_fg") = dy_fg
                        stAdd("wh_receive") = 0
                        stAdd("wh_cancel") = 0
                        stAdd("wh_total") = 0
                        stAdd("return_from_dyeing") = 0
                        stAdd("deli_dy") = 0
                        stAdd("deli_fg") = 0
                        stAdd("deli_sp") = 0
                        stAdd("deli_sa") = 0  '[2022.jan.11] spinning-sample
                        stAdd("deli_total") = 0
                        stAdd("loss") = 0
                        stAdd("remark") = remark
                        stAdd("balance_last_day") = balance_last_day
                        stAdd("balance_today") = balance_last_day
                        stAdd("hk_dmtr") = hk_dmtr
                        dt_last_day.Rows.Add(stAdd)
                    End If
                Next

                'bugs disini, yg datanya dobel harus di gabungin dulu
                For Each rAdd As DataRow In dt_last_day.Rows
                    rAdd("item_code") = rAdd("item_code")
                    If rAdd("balance_last_day") = 0 Then
                        rAdd("balance_today") = (rAdd("wh_receive") - rAdd("wh_cancel")) + rAdd("return_from_dyeing") - (rAdd("deli_dy") + rAdd("deli_fg") + rAdd("deli_sp") + rAdd("deli_sa")) '- rAdd("loss")
                    Else
                        rAdd("balance_today") = rAdd("balance_last_day") + (rAdd("wh_receive") - rAdd("wh_cancel")) + rAdd("return_from_dyeing") - (rAdd("deli_dy") + rAdd("deli_fg") + rAdd("deli_sp") + rAdd("deli_sa")) '- rAdd("loss")
                    End If
                Next

                For st As Integer = 0 To dt_last_day.Rows.Count - 1
                    category = dt_last_day.Rows(st)("category").ToString()
                    item_code = dt_last_day.Rows(st)("item_code").ToString()
                    blended_ratio = dt_last_day.Rows(st)("blended_ratio").ToString()
                    purc_yarn_name = dt_last_day.Rows(st)("purc_yarn_name").ToString()
                    supp_name = dt_last_day.Rows(st)("supp_name").ToString()
                    smm = dt_last_day.Rows(st)("smm").ToString()
                    dmm = dt_last_day.Rows(st)("dmm").ToString()
                    co_hk_pc = dt_last_day.Rows(st)("co_hk_pc").ToString()
                    co_hk_pc_weight = IIf(dt_last_day.Rows(st)("co_hk_pc_weight").ToString() = "", 0, dt_last_day.Rows(st)("co_hk_pc_weight").ToString())
                    dy_fg = dt_last_day.Rows(st)("dy_fg").ToString()
                    wh_receive = dt_last_day.Rows(st)("wh_receive").ToString()
                    wh_cancel = dt_last_day.Rows(st)("wh_cancel").ToString()
                    wh_total = CDec(dt_last_day.Rows(st)("wh_receive").ToString()) - CDec(dt_last_day.Rows(st)("wh_cancel").ToString())
                    return_from_dyeing = dt_last_day.Rows(st)("return_from_dyeing").ToString()
                    deli_dy = dt_last_day.Rows(st)("deli_dy").ToString()
                    deli_fg = dt_last_day.Rows(st)("deli_fg").ToString()
                    deli_sp = dt_last_day.Rows(st)("deli_sp").ToString()
                    deli_sa = dt_last_day.Rows(st)("deli_sa").ToString()
                    deli_total = CDec(dt_last_day.Rows(st)("deli_dy").ToString()) + _
                                 CDec(dt_last_day.Rows(st)("deli_fg").ToString()) + _
                                 CDec(dt_last_day.Rows(st)("deli_sp").ToString()) + _
                                 CDec(dt_last_day.Rows(st)("deli_sa").ToString())
                    remark = dt_last_day.Rows(st)("remark").ToString()
                    loss = dt_last_day.Rows(st)("loss").ToString()
                    balance_last_day = 0
                    balance_today = 0
                    hk_dmtr = IIf(dt_last_day.Rows(st)("hk_dmtr").ToString() = "", 0, dt_last_day.Rows(st)("hk_dmtr").ToString())
                    adaData = False

                    'masukin stocknya ke datalam data di hari kemaren
                    Dim dataLastDay As Boolean = False

                    For Each dbl As DataRow In dt_last_day.Rows
                        If category = dbl("category").ToString And item_code = dbl("item_code").ToString And _
                        blended_ratio = dbl("blended_ratio").ToString And purc_yarn_name = dbl("purc_yarn_name").ToString And _
                        supp_name = dbl("supp_name").ToString And smm = dbl("smm").ToString And _
                        dmm = dbl("dmm").ToString And co_hk_pc = dbl("co_hk_pc").ToString And _
                        co_hk_pc_weight = dbl("co_hk_pc_weight").ToString And dy_fg = dbl("dy_fg").ToString And _
                        remark = dbl("remark").ToString Then
                            balance_last_day += dbl("balance_last_day")
                            balance_today += dbl("balance_today")
                        End If
                    Next

                    'disnininini
                    For Each rowAdd As DataRow In dt_add.Rows
                        If category = rowAdd("category").ToString And item_code = rowAdd("item_code").ToString And blended_ratio = rowAdd("blended_ratio").ToString And purc_yarn_name = rowAdd("purc_yarn_name").ToString And supp_name = rowAdd("supp_name").ToString And smm = rowAdd("smm").ToString And dmm = rowAdd("dmm").ToString And co_hk_pc = rowAdd("co_hk_pc").ToString And co_hk_pc_weight = rowAdd("co_hk_pc_weight").ToString And dy_fg = rowAdd("dy_fg").ToString And remark = rowAdd("remark").ToString Then
                            rowAdd("balance_last_day") = balance_last_day
                            rowAdd("wh_receive") = wh_receive + rowAdd("wh_receive")
                            rowAdd("wh_cancel") = wh_cancel + rowAdd("wh_cancel")
                            rowAdd("wh_total") = rowAdd("wh_receive") - rowAdd("wh_cancel")
                            rowAdd("return_from_dyeing") = return_from_dyeing + rowAdd("return_from_dyeing")
                            rowAdd("deli_dy") = deli_dy + rowAdd("deli_dy")
                            rowAdd("deli_fg") = deli_fg + rowAdd("deli_fg")
                            rowAdd("deli_sp") = deli_sp + rowAdd("deli_sp")
                            rowAdd("deli_sa") = deli_sa + rowAdd("deli_sa")
                            rowAdd("deli_total") = rowAdd("deli_dy") + rowAdd("deli_fg") + rowAdd("deli_sp") + rowAdd("deli_sa")
                            rowAdd("loss") = loss + rowAdd("loss")
                            'rowAdd("balance_today") = balance_last_day + rowAdd("wh_total") + rowAdd("return_from_dyeing") - rowAdd("deli_total") '- rowAdd("loss")
                            rowAdd("balance_today") = balance_last_day + (rowAdd("wh_total")) + rowAdd("return_from_dyeing") - (rowAdd("deli_total")) '- rowAdd("loss")
                            dataLastDay = True
                        Else
                            'rowAdd("return_from_dyeing") = return_from_dyeing
                            rowAdd("balance_today") = rowAdd("wh_total") + rowAdd("return_from_dyeing") - rowAdd("deli_total") '- rowAdd("loss")

                        End If
                    Next

                    'masukin data yang di hari kemaren
                    If dt_add.Rows.Count > 0 Then
                        If dataLastDay = False Then
                            Dim add_last As DataRow = dt_add.NewRow
                            add_last("cur_date") = dt_add.Rows(0)("cur_date").ToString
                            add_last("category") = category
                            add_last("item_code") = item_code
                            add_last("blended_ratio") = blended_ratio
                            add_last("purc_yarn_name") = purc_yarn_name
                            add_last("supp_name") = supp_name
                            add_last("smm") = smm
                            add_last("dmm") = dmm
                            add_last("co_hk_pc") = co_hk_pc
                            add_last("co_hk_pc_weight") = co_hk_pc_weight
                            add_last("dy_fg") = dy_fg
                            add_last("wh_receive") = wh_receive '0
                            add_last("wh_cancel") = wh_cancel '0
                            add_last("wh_total") = wh_total '0
                            add_last("return_from_dyeing") = return_from_dyeing '0
                            add_last("deli_dy") = deli_dy '0
                            add_last("deli_fg") = deli_fg '0
                            add_last("deli_sp") = deli_sp '0
                            add_last("deli_sa") = deli_sa '0
                            add_last("deli_total") = deli_total '0
                            add_last("loss") = loss '0
                            add_last("remark") = remark  'dt_last_day.Rows(st)("remark").ToString
                            add_last("balance_last_day") = balance_last_day
                            add_last("balance_today") = balance_today
                            add_last("hk_dmtr") = hk_dmtr
                            dt_add.Rows.Add(add_last)
                        Else

                        End If
                    End If
                Next

                For Each rFix As DataRow In dt_add.Rows
                    rFix("item_code") = rFix("item_code")
                    rFix("balance_today") = IIf(rFix("balance_last_day").ToString = "", 0, rFix("balance_last_day")) + rFix("wh_total") + rFix("return_from_dyeing") - rFix("deli_total") '- rFix("loss")
                Next

            End If

            dt_add = dt_add
            'balance_last_day - nya "" jadikan 0
            For Each nol As DataRow In dt_add.Rows
                If nol("balance_last_day").ToString = "" Then
                    nol("balance_last_day") = 0
                End If
            Next

            'stock yang sudah habis (balance_last_day <> 0 or balance_today <> 0 ) tidak usah di munculkan
            dt_add.DefaultView.RowFilter = "balance_last_day <> 0 OR wh_total <> 0 OR return_from_dyeing <> 0 OR deli_total <> 0 OR balance_today <> 0 "
            Dim newDt As DataTable = dt_add.DefaultView.ToTable
            ds.Tables.Remove("dt_daily") 'remove dt_daily
            ds.Tables.Add(newDt) 'add dt_daily

            'CEK BALANCE TODAY
            Dim total_balance_today As Decimal = 0
            For Each ttl As DataRow In newDt.Rows
                total_balance_today += ttl("BALANCE_TODAY")
            Next
            total_balance_today = total_balance_today
            kon.Close()
            'tes

            If output = True Then 'SCREEN
                Dim frmReport As frm_report
                frmReport = New frm_report(New rpt_monthly_report_of_raw_white_stock, ds)
                frmReport.Text = "MONTHLY REPORT OF RAW WHITE STOCK| MONTH : " & Format(dtDate.Value, "MMMM yyyy")
                frmReport.ShowDialog()
            Else 'PRINTER
                Dim _repDocGarment As CrystalDecisions.CrystalReports.Engine.ReportDocument
                _repDocGarment = New rpt_monthly_report_of_raw_white_stock
                _repDocGarment.SetDataSource(ds)

                Dim printDialog As New PrintDialog

                If _printName = Nothing Then
                    If printDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        _printName = printDialog.PrinterSettings.PrinterName
                    End If
                End If
                _repDocGarment.PrintOptions.PrinterName = _printName
                _repDocGarment.PrintToPrinter(printDialog.PrinterSettings.Copies, printDialog.PrinterSettings.Collate, _
                  printDialog.PrinterSettings.FromPage, printDialog.PrinterSettings.ToPage)

            End If
        Catch ex As Exception
            Dim Result As String = ""
            Dim st As StackTrace = New StackTrace(ex, True)
            For Each sf As StackFrame In st.GetFrames
                If sf.GetFileLineNumber() > 0 Then
                    If Result = "" Then
                        Result = "Line: " & sf.GetFileLineNumber() & " - Filename: " & IO.Path.GetFileName(sf.GetFileName)
                    Else
                        Result = Result & Environment.NewLine & _
                                 "Line: " & sf.GetFileLineNumber() & " - Filename: " & IO.Path.GetFileName(sf.GetFileName)
                    End If
                End If
            Next
            MessageBox.Show("Error Message: " & ex.Message & Environment.NewLine & Result)
        End Try
    End Sub
End Class