Imports System.Data.SqlClient

Public Class WebForm1
    Inherits System.Web.UI.Page

    ' Connection String สำหรับฐานข้อมูล AdventureWorks2008R2
    Dim connectionString As String = "Data Source=DESKTOP-LJ58AB0\SQLEXPRESS;Initial Catalog=AdventureWorks2008R2;Integrated Security=True;"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ' เรียกเมทอด BindDropDownList เพื่อเตรียมข้อมูลวันที่ใน DropDownList
            BindDropDownList()
        End If
    End Sub

    ' เมทอดสำหรับการเตรียมข้อมูลวันที่ใน DropDownList
    Protected Sub BindDropDownList()
        ' เชื่อมต่อกับฐานข้อมูล
        Using con As New SqlConnection(connectionString)
            con.Open()

            ' สร้างคำสั่ง SQL เพื่อดึงวันที่ที่ไม่ซ้ำกัน
            Dim query As String = "SELECT DISTINCT CONVERT(VARCHAR, OrderDate, 23) AS FormattedOrderDate FROM Purchasing.PurchaseOrderHeader"

            ' ดึงข้อมูลจาก SqlDataReader
            Using cmd As New SqlCommand(query, con)
                ' เตรียมข้อมูลใน DropDownList
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dd_dateinput.DataSource = reader
                    dd_dateinput.DataTextField = "FormattedOrderDate"
                    dd_dateinput.DataValueField = "FormattedOrderDate"
                    dd_dateinput.DataBind()
                End Using
            End Using
        End Using


    End Sub

    Protected Sub submit_Click(sender As Object, e As EventArgs) Handles btn_submit.Click
        ' ... (Your existing code)
        Dim selectedDate As String = dd_dateinput.SelectedValue

        ' สร้าง DataTable เพื่อเก็บข้อมูล
        Dim dt As New DataTable()

        ' เชื่อมต่อกับฐานข้อมูล
        Using con As New SqlConnection(connectionString)
            con.Open()

            ' สร้างคำสั่ง SQL เพื่อดึงรายการสั่งซื้อสำหรับวันที่ที่เลือก
            Dim query As String = "SELECT Poh.PurchaseOrderID, Poh.EmployeeID, Poh.VendorID, Pod.OrderQty, Poh.OrderDate 
                        FROM Purchasing.PurchaseOrderHeader Poh 
                        JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID 
                       WHERE CONVERT(VARCHAR, OrderDate, 23) = @SelectedDate"

            ' ดำเนินการ Execute SQL command
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@SelectedDate", selectedDate)

                ' สร้าง DataAdapter เพื่อเตรียมข้อมูลใน DataTable
                Using da As New SqlDataAdapter(cmd)
                    ' เตรียมข้อมูลใน DataTable
                    da.Fill(dt)
                End Using
            End Using

            ' สร้างคำสั่ง SQL เพื่อดึงผลรวมของ OrderQty สำหรับวันที่ที่เลือก
            Dim querysum_Qty As String = "SELECT ISNULL(SUM(Pod.OrderQty), 0) AS TotalOrderQty " &
                                      "FROM Purchasing.PurchaseOrderHeader Poh " &
                                      "JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID " &
                                      "WHERE CONVERT(DATE, Poh.OrderDate) = @SelectedDate"

            ' ดำเนินการ Execute SQL command เพื่อรับผลรวมของ OrderQty
            Using cmdSumQty As New SqlCommand(querysum_Qty, con)
                cmdSumQty.Parameters.AddWithValue("@SelectedDate", selectedDate)

                ' ดึงผลรวมของ OrderQty
                Dim totalOrderQty As Integer = Convert.ToInt32(cmdSumQty.ExecuteScalar())

                ' แสดงจำนวนรายการและผลรวม Qty
                lbl_listcount.Text = "รวมป้อน " & dt.Rows.Count.ToString() & " รายการ"
                lbl_sumqty.Text = "รวม Qty = " & totalOrderQty.ToString()
            End Using
        End Using

        ' และเอา DataTable เข้ากับ GridView
        GridView1.DataSourceID = "" ' Set DataSourceId to an empty string
        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub


End Class
