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

    ' เมทอดที่เรียกเมื่อคลิกปุ่ม Submit
    Protected Sub submit_Click(sender As Object, e As EventArgs) Handles submit.Click
        ' ดึงวันที่ที่เลือกจาก DropDownList
        Dim selectedDate As String = dd_dateinput.SelectedValue

        ' เชื่อมต่อกับฐานข้อมูล
        Using con As New SqlConnection(connectionString)
            con.Open()

            ' สร้างคำสั่ง SQL เพื่อดึงรายการสั่งซื้อสำหรับวันที่ที่เลือก
            Dim query As String = "SELECT Poh.PurchaseOrderID, Poh.EmployeeID, Poh.VendorID, Pod.OrderQty, Poh.OrderDate 
                                     FROM Purchasing.PurchaseOrderHeader Poh 
                                     JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID 
                                    WHERE CONVERT(VARCHAR, OrderDate, 23) = @SelectedDate"

            ' สร้างคำสั่ง SQL เพื่อดึงผลรวมของ OrderQty สำหรับวันที่ที่เลือก
            Dim querysum_Qty As String = "SELECT ISNULL(SUM(Pod.OrderQty), 0) AS TotalOrderQty " &
                                          "FROM Purchasing.PurchaseOrderHeader Poh " &
                                          "JOIN Purchasing.PurchaseOrderDetail Pod ON Poh.PurchaseOrderID = Pod.PurchaseOrderID " &
                                          "WHERE CONVERT(DATE, Poh.OrderDate) = @SelectedDate"

            ' ดำเนินการ Execute SQL commands
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@SelectedDate", selectedDate)

                ' Execute SQL command และเตรียมข้อมูลใน GridView
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    GridView1.DataSource = reader
                    GridView1.DataBind()
                End Using
            End Using

            ' ดำเนินการ Execute SQL command เพื่อรับผลรวมของ OrderQty
            Using cmdSumQty As New SqlCommand(querysum_Qty, con)
                cmdSumQty.Parameters.AddWithValue("@SelectedDate", selectedDate)

                ' ดึงผลรวมของ OrderQty
                Dim totalOrderQty As Integer = Convert.ToInt32(cmdSumQty.ExecuteScalar())

                ' แสดงจำนวนรายการและผลรวม Qty
                lbl_list_count.Text = "รวมป้อน " & GridView1.Rows.Count.ToString() & " รายการ"
                lbl_sum_qty.Text = "รวม Qty = " & totalOrderQty.ToString()
            End Using
        End Using
    End Sub

End Class
