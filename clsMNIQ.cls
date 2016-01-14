VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMNIQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

Private mvaruserid As String
Private mvarDepthTill As Long
Private mvarDepthNow As Long
Private mvarSubject As String

Private mvarOrder As Long
Private mvarMaxPos As Long
Private mvarGeneralHang As Long
Private mvarQuizNum As Long

'Private Function getNewH1(ByRef O_ As Long, ByVal order_new As Long, ByVal depth1_old As Long, ByVal depth2_old As Long) As Long
'Dim i As Long
'Dim h0 As Long
'Dim depth1_new As Long
'Dim h As Long
'Dim sumVal As Long
'
'sumVal = 0
'h0 = geth0(O_, depth1_new, order_new)
'i = 1
'
'Do While (i <= depth2_old)
'    sumVal = sumVal + order_new ^ i
'    i = i + 1
'Loop
'
'getNewH1 = sumVal
'
'End Function

'================================================================================================
'O_�� order �� ���� h0 �� depth1�� ���Ѵ�.
'return h0
'================================================================================================
'Function geth0(ByVal O_ As Long, ByRef depth1 As Long, depth1 As Long, Optional ByVal order As Long = 2, Optional ByVal step As Long = 2) As Long
Public Function geth0(ByVal O_ As Long, ByRef depth1 As Long, ByVal Order As Long) As Long

'���ڰ� �ʹ� ū��� �����߻� ���ɼ� ����
'����ó�� �ڵ鸵 �ʿ���

Dim upper As Long
Dim crossLeft As Long
Dim crossRight As Long
Dim sumVal As Long
Dim h0 As Long

'�ʱ�ȭ
depth1 = 0
sumVal = O_

Debug.Assert Order > 1
While (sumVal \ Order) = (sumVal / Order)
    depth1 = depth1 + 1
    sumVal = sumVal / Order 'depth1�� ������ �ش� O_���� ���������� �ٲ�� �ֵ��� �̷��� ������
Wend

upper = O_ - 1
crossRight = Fix(upper / Order)
'crossLeft = step
crossLeft = Order
sumVal = O_

'O_ -> ����(h0)���ϱ�
Do Until (crossRight <= 0)
    sumVal = sumVal + crossLeft * crossRight
    'crossLeft = crossLeft * step
    crossLeft = crossLeft * Order
    crossRight = Fix(crossRight / Order)
'    depth1 = depth1 + 1
Loop
geth0 = sumVal

End Function

Private Function isOdd(ByVal val As Long) As Boolean
isOdd = (Fix(Trim(val) / 2) <> (val / 2))
End Function
'Public Function calQuizNum1() As Long
'Dim isNew As Boolean
'calQuizNum1 = getF(mvarMaxPos, mvarGeneralHang, mvarOrder, isNew, mvarDepthTill, mvarDepthNow)
'
'End Function


Public Function getF(ByRef O_ As Long, ByRef H As Long, ByVal Order As Long, ByRef isNew As Boolean, ByRef depth1 As Long, ByRef depth2 As Long, Optional ByRef ph0 As Long, Optional ByRef pH1 As Long) As Long
Dim h0 As Long
Dim h1 As Long

'Dim depth1 As Long
'Dim depth2 As Long

If Order < 2 Then
    err.Raise vbObjectError + 1052, "order�� ���� 2�̻��̾�� �մϴ�!"
End If

'---------------------------------------------------------------------------
'Test �� O_ h���� ����
'---------------------------------------------------------------------------
Dim TempO_ As Long
Dim Temph0 As Long

If O_ < 1 Then
    O_ = 1 'O_ �� �ּҰ� ���� ������ �ڵ� ����
End If

TempO_ = O_ + 1
Temph0 = geth0(TempO_, depth1, Order)

isNew = False
If H = Temph0 Then
    O_ = TempO_ 'O_���� ���� ���� ����
    isNew = True
ElseIf H > Temph0 Then
    H = Temph0 - 1 'h�� �ִ밪 ���� ���� ����
'    Debug.Assert False
Else
    '�ڷΰ��⸦ �Ѱ��
    
End If

Temph0 = geth0(O_, depth1, Order)
If H <= Temph0 Then
    If (H + 1) = Temph0 Then
        O_ = O_ - 1
        If O_ < 1 Then
            O_ = 1
            H = 1
        End If
    Else
        H = Temph0 'h�� �ּҰ� ���� ���� ����
        isNew = True
    End If 'Debug.Assert False
End If
'---------------------------------------------------------------------------

h1 = getH1(O_, H, depth2, Order)

getF = H - h1 + O_
ph0 = H
pH1 = h1
End Function

Private Function getH1(ByVal O_ As Long, ByVal H As Long, ByRef depth2 As Long, ByVal Order As Long) As Long

Dim i As Long
Dim depth1 As Long
Dim sumVal As Long
Dim h0 As Long

h0 = geth0(O_, depth1, Order)

sumVal = h0
depth2 = 0
i = 1
Do Until sumVal >= H
    sumVal = sumVal + Order ^ i
    depth2 = depth2 + 1
    i = i + 1
Loop

getH1 = sumVal

End Function

