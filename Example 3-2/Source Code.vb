Imports System.IO
Imports System.IO.IOException

Public Class Form1
    Inherits System.Windows.Forms.Form

    'Declare All Variables in Program
    Dim position, num, unit, money_now, sale_price, total_price As Integer
    Dim name As String

    'Proucdure To Clear All TextBoxes
    Sub clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox1.Focus()
    End Sub

    'Proucdure To Display Data in ListBox
    Sub view()

        FileOpen(1, "xx.txt", OpenMode.Input)

        ListBox1.Height = (ListBox1.ItemHeight + 3) * 12
        ListBox1.Items.Clear() 'لمسح البيانات الموجوده في الليست بوكس

        ListBox1.Items.Add("") 'علشان نسيب اول سطر ف الليست بوكس فاضي ويبدأ يكتب من السطر التاني
        ListBox1.Items.Add("|" + "رقم الصف")
        ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
        ListBox1.Items.Add("|" + "اسم الصنف")
        ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
        ListBox1.Items.Add("|" + "الوحده")
        ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
        ListBox1.Items.Add("|" + "الرصيد الحالي")
        ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
        ListBox1.Items.Add("|" + "السعر")
        ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
        ListBox1.Items.Add("|" + " السعر الكلي")

        Do While Not EOF(1)

            'بنجيب بيانات أول صف ف الفايل
            Input(1, num)
            Input(1, name)
            Input(1, unit)
            Input(1, money_now)
            Input(1, sale_price)
            Input(1, total_price)

            'بنعرض البيانات في الليست بوكس
            ListBox1.Items.Add("")
            ListBox1.Items.Add("|" + Str(num))
            ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
            ListBox1.Items.Add("|" + name)
            ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
            ListBox1.Items.Add("|" + Str(unit))
            ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
            ListBox1.Items.Add("|" + Str(money_now))
            ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
            ListBox1.Items.Add("|" + Str(sale_price))
            ListBox1.Items.Add("|" + "--------------------------------------------------------------------------")
            ListBox1.Items.Add("|" + Str(total_price))

            'علشان لما اكتب رقم الصنف واضغط علي عرض الصنف يعرضلي كل البيانات بتاعته ف كل التيكست بوكسيز
            If num = Val(TextBox1.Text) Then
                TextBox2.Text = name
                TextBox3.Text = unit
                TextBox4.Text = money_now
                TextBox5.Text = sale_price
                TextBox6.Text = total_price
            End If

        Loop

        FileClose(1)

    End Sub

    'Exit Button
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        End
    End Sub

    'View Button
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        view()
    End Sub

    'Add Button
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        FileOpen(1, "XX.txt", OpenMode.Append)

        num = Val(TextBox1.Text)
        name = TextBox2.Text
        unit = Val(TextBox3.Text)
        money_now = Val(TextBox4.Text)
        sale_price = Val(TextBox5.Text)
        total_price = Val(TextBox6.Text)

        WriteLine(1, num, name, unit, money_now, sale_price, total_price)

        FileClose(1)

        MsgBox("تم حفظ الصنف بنجاح")
        clear()

    End Sub

    'Edit Button
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FileOpen(1, "XX.txt", OpenMode.Input)
        FileOpen(2, "YY.txt", OpenMode.Append)

        Do While Not EOF(1)

            Input(1, num)
            Input(1, name)
            Input(1, unit)
            Input(1, money_now)
            Input(1, sale_price)
            Input(1, total_price)

            'IF this are The Record You want To Edit
            If num = Val(TextBox1.Text) Then
                position = Loc(1)

                name = TextBox2.Text
                unit = Val(TextBox3.Text)
                money_now = Val(TextBox4.Text)
                sale_price = Val(TextBox5.Text)
                total_price = Val(TextBox6.Text)

                If (Loc(1) <> position - 1) Then
                    WriteLine(2, num, name, unit, money_now, sale_price, total_price)
                End If
            End If

            'IF this are Not The Record You want To Edit
            If num <> Val(TextBox1.Text) Then
                position = Loc(1)
                If (Loc(1) <> position - 1) Then
                    WriteLine(2, num, name, unit, money_now, sale_price, total_price)
                End If
            End If

        Loop

        FileClose(2)
        FileClose(1)

        Kill("XX.txt")
        Rename("YY.txt", "XX.txt")
        clear()
    End Sub

    'Delete Button
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        FileOpen(1, "XX.txt", OpenMode.Input)
        FileOpen(2, "YY.txt", OpenMode.Append)

        Do While Not EOF(1)

            Input(1, num)
            Input(1, name)
            Input(1, unit)
            Input(1, money_now)
            Input(1, sale_price)
            Input(1, total_price)

            'IF this are Not The Record You want To Delete
            If num <> Val(TextBox1.Text) Then

                position = Loc(1)
                If (Loc(1) <> position - 1) Then
                    WriteLine(2, num, name, unit, money_now, sale_price, total_price)
                End If
            End If

        Loop

        FileClose(2)
        FileClose(1)

        Kill("XX.txt")
        Rename("YY.txt", "XX.txt")
        clear()

    End Sub

    'استحدمنا حدث آخر لزرار اضافه عنصر جديد علشان بمجرد م اضغط عليه واضيف عنصر جديدالتعديل يظهر ف الليست بوكس
    Private Sub Button1_MouseUp(sender As Object, e As MouseEventArgs) Handles Button1.MouseUp
        view()
    End Sub

    'استحدمنا حدث آخر لزرار التعديل علشان بمجرد م اضغط عليه واعدل التعديل يظهر ف الليست بوكس
    Private Sub Button3_MouseUp(sender As Object, e As MouseEventArgs) Handles Button3.MouseUp
        view()
    End Sub

    'استحدمنا حدث آخر لزرار حذف عنصر علشان بمجرد م اضغط عليه واحذف عنصر التعديل يظهر ف الليست بوكس
    Private Sub Button4_MouseUp(sender As Object, e As MouseEventArgs) Handles Button4.MouseUp
        view()
    End Sub

End Class
