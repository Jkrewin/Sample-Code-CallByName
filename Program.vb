
Module Program
    Sub Main(args As String())
        Console.Title = "����������� �����"

        ' ��������� ��� ������ ������ - System does not support 'windows-1251' 
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

        REM ��������, ����� �� ���������
        AddHandler Console.CancelKeyPress, AddressOf KeyHandl

        REM ����� ���������
        Console.Write("����������� �����   ") : Console.ForegroundColor = ConsoleColor.Blue : Console.WriteLine("{����� �� ��������� Cntr+C}")
        Console.ForegroundColor = ConsoleColor.White

        Dim sector As New myCommand

        REM ����� �������� ���� BASIC 
        On Error GoTo ErrorHandler
start:
        Console.Write("������� ������� ::")
        Dim vega() As String = Split(Console.ReadLine, " ")
        Dim param(vega.Count - 1, 1) As String

        For i = 1 To vega.Count - 1
            param(i, 0) = vega(i).Split("=")(0).Trim
            param(i, 1) = vega(i).Split("=")(1).Trim
        Next

        CallByName(sector, vega(0), CallType.Method, param)
        GoTo start
ErrorHandler:
        Console.WriteLine("�� ������� �� ������ �������!")
        GoTo start
    End Sub

    Sub KeyHandl(sender As Object, e As ConsoleCancelEventArgs)
        Console.WriteLine("����������� ���������.")
        Environment.Exit(0)
    End Sub
End Module

''' <summary>
''' ��� ��� ������� 
''' </summary>
Public Class myCommand
    Private currencies As New List(Of String) From {"USD", "RUB", "EUR"} ' ��������� ����� ������ �������� ������
    Private Wallet As New Dictionary(Of String, Decimal)
    Private Course As New List(Of Valute)

    Public Sub New()
        Print = "                         =    ������ �������  ="
        For Each tv In Me.GetType.GetMethods
            If tv.GetCustomAttributesData.Count <> 0 AndAlso "Information" = tv.GetCustomAttributesData(0).AttributeType.Name Then
                Print = tv.Name & " :: " & tv.GetCustomAttributesData(0).ConstructorArguments(0).Value : Print = ""
            End If
        Next
        Print = "-----------------------------------------------------------------------------------"
        '�������� ��� ��� ������ ��� ����� ����� �� ��������� �� null � �������
        currencies.ForEach(Sub(x) Wallet.Add(x, 0))
        ' ��������� ����������
        Check(Nothing)
    End Sub

    <Information("������� �����")>
    Public Sub Cls(args(,) As String)
        Console.Clear()
    End Sub

    <Information("��������� ��� ����������� ������. Valuta [check=all-���������� ������ (all ����� �������� �� ������)] [set=��������� ����� ������] [del=������� ������] ������: valuta set=CNY ��� valuta check set=CNY")>
    Public Sub Valuta(args(,) As String)
        For i = 1 To UBound(args) - LBound(args)
            Dim p As String = Mid(args(i, 1), 1, 3).ToUpper
            Select Case args(i, 0).ToLower
                Case "check"
                    If Course.Exists(Function(x) x.Show(p)) = False Then
                        currencies.ForEach(Sub(x) Console.WriteLine("[" & x & "] "))
                    End If
                Case "set"
                    If currencies.Contains(p) = True Then
                        Print = "����� ������ ��� ���������� � ������."
                    ElseIf Course.Exists(Function(x) x.CharCode = p) = False Then
                        Print = "����� ������ �� ����������."
                    Else
                        currencies.Add(p)
                        Wallet.Add(p, 0)
                    End If
                Case "del"
                    Dim t As Integer = currencies.FindIndex(Function(x) x = p)
                    If i = -1 Then
                        Print = "��� ����� ��������: " & p
                    Else
                        currencies.RemoveAt(t)
                    End If
                Case Else
                    Print = "������� �������� ��������� ������� !"
            End Select
        Next
    End Sub


    <Information("��������� ������. FullUpWallet [��� ������ USD, RUB, EUR]=[�����] ������: fullupwallet usd=150")>
    Public Sub FullUpWallet(args(,) As String)
        If currencies.Exists(Function(x) x = args(1, 0).ToUpper) = False Or IsNumeric(args(1, 1)) = False Then
            Print = "����� ������ �� ����������, ���� �� ������� ������� ����� !"
        Else
            Wallet(args(1, 0).ToUpper) += CDec(args(1, 1))
        End If
    End Sub

    <Information("������� ������ �������.")>
    Public Sub List(args(,) As String)
        For Each tv In Wallet
            Console.WriteLine($"        {tv.Key}= {tv.Value }")
        Next
    End Sub

    <Information("�������� ������ �� ��������� ����� � ������� ����. �������� � �������� ������� ������ �� cbr.ru")>
    Public Sub Check(args(,) As String)
        Dim Url As String = "https://cbr.ru/scripts/XML_daily.asp?date_req=" & String.Format("{0:0#}", Now.Day) & "/" & String.Format("{0:0#}", Now.Month) & "/" & Now.Year
        Dim doc As XDocument

        Print = "������� �������� � ������ ����� �� cbr.ru"
        Try
            Using oWeb = New Net.WebClient()
                Using dat As IO.Stream = oWeb.OpenRead(Url)
                    doc = XDocument.Load(dat)
                End Using
            End Using
        Catch
            Print = "��� ������� � cbr.ru"
            Exit Sub
        End Try

        Dim result = From tv In doc.Descendants("Valute") Select New Valute With {
                                                             .CharCode = tv.Elements("CharCode").Value,
                                                             .Value = tv.Elements("Value").Value,
                                                             .Name = tv.Elements("Name").Value,
                                                             .Nominal = tv.Elements("Nominal").Value,
                                                             .NumCode = tv.Elements("NumCode").Value}
        Course.Clear()
        Console.ForegroundColor = ConsoleColor.Blue
        For Each tv As Valute In result
            Course.Add(tv)
            If currencies.Contains(tv.CharCode) Then Console.WriteLine(tv.ToString)
        Next
        Console.ForegroundColor = ConsoleColor.White
    End Sub

    <Information("������������ ������. Convert [��� ������ USD, RUB, EUR]>=[��� ������ USD, RUB, EUR] sum=[����� � ������]. ������ Convert USD>=RUB sum=10")>
    Public Sub Convert(args(,) As String)
        If Course.Count = 0 Then Print = "�������� �������� check ��������" : Exit Sub

        Dim fn = Function(tv As String) As Decimal
                     If tv = "RUB" Then Return 1
                     Return Course.Find(Function(x) x.CharCode = tv).Value
                 End Function
        Try
            Dim A, B As String
            A = Microsoft.VisualBasic.Left(args(1, 0).ToUpper, 3)
            B = args(1, 1).ToUpper

            If currencies.Exists(Function(x) x = A) = False Or currencies.Exists(Function(x) x = B) = False Then
                Print = "����� ������ �� ���������� !"
            Else
                Dim d As Decimal = args(2, 1)                     '��� ��������
                Dim money As Decimal = Wallet(A)         '������� ������

                If d > money Then Print = "��������� ����� ������ ��� ����� ������� �� ����� ��������. ������ ���: " & money : Exit Sub
                Wallet(A) = FormatCurrency(money - d)
                Dim rus As Decimal = d * fn(A) ' ������� � ����� �������
                Wallet(B) += FormatCurrency(IIf(B = "RUB", rus, rus / fn(B)))
                List(Nothing)
                Exit Sub
            End If
        Catch
            Print = "������� ������� ����� ��� ������� �������� ��������� ������� !"
        End Try
    End Sub

    Private WriteOnly Property Print As String
        Set(value As String)
            Dim c As ConsoleColor = Console.ForegroundColor
            If value = "" Then
            ElseIf Microsoft.VisualBasic.Right(value, 1) = "!" Then
                Console.ForegroundColor = ConsoleColor.Red
            End If
            Console.WriteLine(value)
            Console.ForegroundColor = c
        End Set
    End Property


    Private Structure Valute
        Public NumCode As Short
        Public CharCode As String
        Public Nominal As Integer
        Public Name As String
        Public Value As String

        Public Shadows ReadOnly Property ToString As String
            Get
                Return "     " & Name & vbTab & Value * Nominal & " (" & CharCode & ")"
            End Get
        End Property

        ''' <summary>
        ''' ��������� � ����� ������� �� ������� ���� ���������� 
        ''' </summary>
        ''' <param name="a"></param>
        ''' <returns></returns>
        Public Function Show(a As String) As Boolean
            If a.ToUpper = CharCode Then
                Console.WriteLine($" NumCode= {NumCode}; CharCode= {CharCode}; Nominal={Nominal}; Name= {Name}; Value= {Value}")
                Return True
            Else
                Return False
            End If
        End Function
    End Structure
    Private Class Information
        Inherits Attribute
        Private _name As String

        Public Sub New(ByVal name As String)
            _name = name
        End Sub
        Public ReadOnly Property Name() As String
            Get
                Return _name
            End Get
        End Property
    End Class
End Class