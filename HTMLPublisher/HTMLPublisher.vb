'***********************************************************************
'Copyright 2005-2022 Vijay Sridhara

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at

'http://www.apache.org/licenses/LICENSE-2.0

'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'***********************************************************************
''' <summary>
''' HTML Publisher can publish your reports
''' </summary>
''' <remarks>
''' The code is licensed under Microsoft Public license. You must have received a copy
''' of the license along with the code. Please write to vitallogic@gmail.com for any questions. 
''' The code is copyrighted 2008-2011 by Vijay Sridhara. 
''' </remarks>

Public Class HTMLPublisher


    Private HFont As Font = New Font("Verdana", 2, FontStyle.Bold, GraphicsUnit.World)
    Public Property HeaderFont() As Font
        Get
            Return HFont
        End Get
        Set(ByVal value As Font)
            HFont = value
            HBold = HFont.Bold
        End Set
    End Property
    Private HBold As Boolean = True
    Private DBold As Boolean = False
    Private DFont As Font = New Font("Verdana", 2, FontStyle.Regular, GraphicsUnit.World)
    Public Property DataFont() As Font
        Get
            Return DFont
        End Get
        Set(ByVal value As Font)
            DFont = value
            DBold = DFont.Bold
        End Set
    End Property
    Private HNeeded As Boolean = True
    Public Property HeaderNeeded() As Boolean
        Get
            Return HNeeded
        End Get
        Set(ByVal value As Boolean)
            HNeeded = value
        End Set
    End Property
    Private FNeeded As Boolean = True
    Public Property FooterNeeded() As Boolean
        Get
            Return FNeeded
        End Get
        Set(ByVal value As Boolean)
            FNeeded = value
        End Set
    End Property
    Private DocBackColor As Color = Color.Beige
    Public Property DocumentBackColor() As Color
        Get
            Return DocBackColor
        End Get
        Set(ByVal value As Color)
            DocBackColor = value
        End Set
    End Property
    Private HColor As Color = Color.Black
    Public Property HeaderColor() As Color
        Get
            Return HColor
        End Get
        Set(ByVal value As Color)
            HColor = value
        End Set
    End Property
    Private HBColor As Color = Color.LightGray
    Public Property HeaderBackColor() As Color
        Get
            Return HBColor
        End Get
        Set(ByVal value As Color)
            HBColor = value
        End Set
    End Property
    Private DColor As Color = Color.Black
    Public Property DataColor() As Color
        Get
            Return DColor
        End Get
        Set(ByVal value As Color)
            DColor = value
        End Set
    End Property
    Private TBClolor As Color = Color.Khaki
    Public Property TableBackColor() As Color
        Get
            Return TBClolor
        End Get
        Set(ByVal value As Color)
            TBClolor = value
        End Set
    End Property
    Private TBorder As Integer = 0
    Public Property TableBorder() As Integer
        Get
            Return TBorder
        End Get
        Set(ByVal value As Integer)
            TBorder = value
        End Set
    End Property
    Private CSpacing As Integer = 2
    Public Property CellSpacing() As Integer
        Get
            Return CSpacing
        End Get
        Set(ByVal value As Integer)
            CSpacing = value
        End Set
    End Property
    Private Header As String = "Created by HTML Publisher on " & Format(Date.Now, "dd-MMM-yyyy")
    Public Property HeaderText() As String
        Get

            Return Header
        End Get
        Set(ByVal value As String)
            Header = value
        End Set
    End Property
    Private Footer As String = "End of Report"
    Public Property FooterText() As String
        Get
            Return Footer
        End Get
        Set(ByVal value As String)
            Footer = value
        End Set
    End Property
    Private _LineNumbering As Boolean
    Public Property LineNumbering() As Boolean
        Get
            Return _LineNumbering
        End Get
        Set(ByVal value As Boolean)
            _LineNumbering = value
        End Set
    End Property
    Private _ColumnHeaders As New Collection
    Private _Data As New List(Of DataRow)
    ''' <summary>
    ''' Adds a data row
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DataRow
        Private _cols As New Collection
        Public ReadOnly Property Fields() As Collection
            Get
                Return _cols
            End Get
        End Property
    End Class
    Public Sub AddHeader(ByVal str As String)
        _ColumnHeaders.Add(str)
    End Sub
    Public ReadOnly Property ColumnHeaders() As Collection
        Get
            Return _ColumnHeaders
        End Get
    End Property
    Public Sub AddRow(ByVal drow As DataRow)
        _Data.Add(drow)
    End Sub
    Public ReadOnly Property Data() As List(Of DataRow)
        Get
            Return _Data
        End Get
    End Property
    Public Enum ECodes
        Success
        DatacolumnsMismatch
        DataEmpty
        HeadersEmpty
        FileWriteError
    End Enum
    Private HTML = "<HTML><TITLE>{0}</TITLE><BODY BGCOLOR={1}>{2}</BODY></HTML>"
    Private TABLE = "<TABLE ALIGN=CENTER WIDTH=100% BORDER={0} CELLSPACING={1}>{2}</TABLE> " & vbCrLf
    Private TABLE_ROW = "<TR>{0}</TR> " & vbCrLf
    Private TABLE_CELL = "<TD BGCOLOR={0}><FONT FACE={1} SIZE={2} COLOR={3}>{4}</FONT></TD> " & vbCrLf
    Private TABLE_HDR = "<TD BGCOLOR={0}><FONT FACE={1} SIZE={2} COLOR={3}>{4}</FONT></TD> " & vbCrLf
    Private HDRFTRTEXT = "<TD BGCOLOR={0} COLSPAN={1} ALIGN=CENTER><FONT FACE={2} SIZE={3} COLOR={4}>{5}</FONT></TD> " & vbCrLf

    Public Function PublishFile(ByVal fpath As String) As ECodes
        Try
            Dim colcnt As Integer = ColumnHeaders.Count
            If colcnt = 0 Then Return ECodes.HeadersEmpty
            Dim rowCnt As Integer = Data.Count
            If rowCnt = 0 Then Return ECodes.DataEmpty
            For i As Integer = 0 To Data.Count - 1
                If Data(i).Fields.Count <> colcnt Then Return ECodes.DatacolumnsMismatch
            Next
            Dim TableContent As String = ""
            ' If LineNumbering Then colcnt += 1
            Dim ColSpan As Integer = colcnt
            If LineNumbering Then ColSpan += 1
            Dim ColStart As Integer = 0
            If LineNumbering Then ColStart = -1
            If HeaderNeeded Then
                If HBold Then
                    TableContent += String.Format(HDRFTRTEXT, GetHex(HeaderBackColor), ColSpan, HeaderFont.Name, HeaderFont.Size, GetHex(HeaderColor), "<B>" & HeaderText & "</B>")
                Else
                    TableContent += String.Format(HDRFTRTEXT, GetHex(HeaderBackColor), ColSpan, HeaderFont.Name, HeaderFont.Size, GetHex(HeaderColor), HeaderText)
                End If
            End If
            TableContent = String.Format(TABLE_ROW, TableContent)
            Dim CellData As String = ""

            For i As Integer = ColStart To colcnt - 1
                If i = ColStart And LineNumbering Then
                    If HBold Then
                        CellData += String.Format(TABLE_CELL, GetHex(HeaderBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), "<B>Sl. No</B>")
                    Else
                        CellData += String.Format(TABLE_CELL, GetHex(HeaderBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), "Sl. No")
                    End If
                Else
                    If HBold Then
                        CellData += String.Format(TABLE_CELL, GetHex(HeaderBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), "<B>" & ColumnHeaders(i + 1).ToString & "</B>")
                    Else
                        CellData += String.Format(TABLE_CELL, GetHex(HeaderBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), ColumnHeaders(i + 1).ToString)
                    End If

                End If
            Next
            ColStart = 1
            If LineNumbering Then ColStart = 0
            TableContent += String.Format(TABLE_ROW, CellData)
            For j As Integer = 0 To rowCnt - 1
                CellData = ""
                For i As Integer = ColStart To colcnt
                    If i = ColStart And LineNumbering Then
                        If HBold Then
                            CellData += String.Format(TABLE_CELL, GetHex(TableBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), "<B>" & j + 1 & "</B>")
                        Else
                            CellData += String.Format(TABLE_CELL, GetHex(TableBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), j + 1)
                        End If
                    Else
                        If DBold Then
                            CellData += String.Format(TABLE_CELL, GetHex(TableBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), "<B>" & Data(j).Fields(i) & "</B>")
                        Else
                            CellData += String.Format(TABLE_CELL, GetHex(TableBackColor), DataFont.Name, DataFont.Size, GetHex(DataColor), Data(j).Fields(i))
                        End If
                    End If
                Next
                TableContent += String.Format(TABLE_ROW, CellData)
            Next

            If FooterNeeded Then
                If HBold Then
                    TableContent += String.Format(HDRFTRTEXT, GetHex(HeaderBackColor), ColSpan, HeaderFont.Name, HeaderFont.Size, GetHex(HeaderColor), "<B>" & FooterText & "</B>")
                Else

                    TableContent += String.Format(HDRFTRTEXT, GetHex(HeaderBackColor), ColSpan, HeaderFont.Name, HeaderFont.Size, GetHex(HeaderColor), FooterText)
                End If
            End If
            TableContent = String.Format(TABLE, TableBorder, CellSpacing, TableContent)
            Dim HTMLContent As String = String.Format(HTML, "Report", GetHex(DocBackColor), TableContent)
            Dim sw As New IO.StreamWriter(fpath)
            sw.Write(HTMLContent)
            sw.Close()
            sw.Dispose()

            Return ECodes.Success
        Catch ex As Exception
            MsgBox(ex.Message)
            Return ECodes.FileWriteError
        End Try
    End Function
    Private Function GetHex(ByVal nu As Color) As String
        Dim fstring As String = "#"
        Dim st As String = Hex(nu.ToArgb)
        Return "#" & st.Substring(2)
    End Function
End Class
