VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QuadratureForm 
   Caption         =   "Numerical Integration (Quadrature) By Romberg Method"
   ClientHeight    =   5190
   ClientLeft      =   1050
   ClientTop       =   330
   ClientWidth     =   8010
   OleObjectBlob   =   "QuadratureForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QuadratureForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'
'   Copyright (C) 2001, 2007  Prof. Jayanth R. Varma, jrvarma@iimahd.ernet.in,
'   Indian Institute of Management, Ahmedabad 380 015, INDIA
'
'   This program is free software; you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation; either version 2 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program (see file COPYING); if not, write to the
'   Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'   Boston, MA  02111-1307  USA
'
'
'This uses the Romberg method to perform numerical integration
'
'
'Global Variables
'xCell and fCell are the cells containing x and f(x)
Dim xCell As Range, fCell As Range, outCell As Range
'a and b define the range of x ( a <= x <= b)
'Eta and Eps are tolerance parameters
Dim a As Double, b As Double, Eta As Double, Eps As Double
'Min number of Romberg iterations
Dim min As Integer, max As Integer
'The computed value of the integral
Dim integral As Double
Dim integral_f As String
'
'
Private Function f(x As Double) As Double
'This converts the spreadsheet into a function f(x)
'The input value is entered into the spreadsheet (xCell)
'and the value f(x) is read from the spreadsheet (fCell)
On Error GoTo function_error
xCell.Value = x
f = fCell.Value
Exit Function
function_error:
MsgBox "Error while trying to evaluate function for x = " & _
       Str$(x)
Unload Me
End Function
'
'
Private Sub integrate()
'Main integration method using Romberg method
Dim q(1 To 100) As Double
' a and b are the limits of integration
' f is the function to be integrated (see function definition at the top)
' Trapezoidal rule with 2^k subintervals of length (a-b)/2^k is
' T(k) = h[f(a)/2 + f(a+h) + f(a+2h) + ... + f(b)/2]
' This can be computed more easily using the recursive definition
' T(k) = [T(k-1) + M(k-1)]/2 where M(k) denotes the mid-point rule
' Consider the triangular array
' T(0,0)
' T(1,0)    T(0,1)
' T(2,0)    T(1,1)  T(0,2)
'   ...
' where T(k,0) denotes T(k) and
' T(k,m) = [4^m T(k+1,m-1) - T(k,m-1)]/(4^m-1)
' The successive Romberg approximants are the diagonal elements of this array
StatusL.Caption = "Iterating"
QuadratureForm.Repaint
H = b - a
fa = f(a)
fb = f(b)
tabs = Abs(H) * (Abs(fa) + Abs(fb)) / 2
' t = T(0)
t = H * (fa + fb) / 2
nx = 1
For n = 1 To max
    H = H / 2
    Sum = 0
    sumabs = 0
    For i = 1 To nx
        xi = 2 * i - 1
        fi = f(a + xi * H)
        sumabs = sumabs + Abs(fi)
        Sum = Sum + fi
    Next i
    ' 2h*Sum is M(n-1) and we compute T(n) from T(n-1) and M(n-1)
    t = t / 2 + H * Sum
    tabs = tabs / 2 + Abs(H) * sumabs
    ' Compute q(n) = T(n-1,1)
    q(n) = 2 * (t + H * Sum) / 3
    If (n >= 2) Then
        g = 4
        For j = 2 To n
            i = n + 1 - j
            g = g * 4
            ' At the end of each iteration:
            ' q is the n'th row of the T array in reverse order with the
            ' element T(n,0) omitted
            ' q(1) = T(0,n);  q(n) = T(n-1,1) and the elements in between are
            ' T(1,n-1), T(2,n-2), T(3,n-3) ...
            ' We now compute the new q array using the previous q array
            ' proceeding from q(n) to q(1)
            q(i) = q(i + 1) + (q(i + 1) - q(i)) / (g - 1)
        Next j
    End If
    If (n >= 3) Then
        ' For convergence tests, we compute the differences between the q(1)
        ' values from the last three iterations
        x = Abs(q(1) - qx2) + Abs(qx2 - qx1)
        ' Absolute error test : |q(1)-I| < eps where I is the true integral
        convergedA = x <= 3 * Eps
        ' Relative error test : |q(1)-I|/|I| < eta or more precisely
        ' q(1) = integral of (1+y(x))f(x)dx where |y(x)| < eta
        ' This is relative error only if f(x) does not change sign
        If (tabs <> 0) Then
            convergedR = (x / tabs <= 3 * Eta)
        Else
            convergedR = False
        End If
        integral = q(1)
        IntegralL.Caption = myFormat(integral)
        AbsoluteL.Caption = myFormat(x / 3)
        RelativeL.Caption = myFormat(x / (tabs + 0.000000000000001) / 3)
        IterationL.Caption = "   " & n
        QuadratureForm.Repaint
        If ((n >= min) And (convergedA Or convergedR)) Then Exit For
    End If
    ' We must store the q(1) values of last two iterations
    If (n >= 2) Then qx1 = qx2
    qx2 = q(1)
    nx = nx * 2
Next n
If (convergedA Or convergedR) Then
    StatusL.Caption = "Success"
    integral_f = integral
Else
    StatusL.Caption = "Error"
    MsgBox "Integral failed to converge. "
    integral_f = "=na()"
End If
QuadratureForm.Repaint
End Sub
'
'
Private Sub CancelButton_Click()
Unload Me
End Sub
'
'
Private Sub CloseButton_Click()
outCell.Value = integral_f
Unload Me
End Sub
'
'
Private Sub IntegrateButton_Click()
Set xCell = Range(xCellRE.Value)
Set fCell = Range(fCellRE.Value)
Set outCell = Range(OutcellRE.Value)
a = LowTB.Value
b = UpTB.Value
Eta = Application.max(Abs(etaTB.Value), 0.00000001)
Eps = Abs(epsTB.Value)
min = minTB.Value
If (min < 3) Then min = 3
max = MaxTB.Value
If (max > 20) Then max = 20
etaTB.Value = Eta
epsTB.Value = Eps
minTB.Value = min
MaxTB.Value = max
'We store the current values of all user parameters as the
'new defaults for this Excel session
'This function is in the Romberg module to allow static
'variables to be used
Call QuadratureDefaults(xCell, fCell, outCell, a, b, Eta, Eps, min, max, True)
'Integrate
QuadratureForm.Repaint
Call integrate
End Sub

'
'
Private Sub UserForm_Initialize()
'Read default values. In case this form has been invoked earlier
'in this Excel session, the values are remembered from that invocation
Call QuadratureDefaults(xCell, fCell, outCell, a, b, Eta, _
                        Eps, min, max, False)
'xCell, fCell and outCell are also remembered from earlier
'invocation if any. Else these are all empty
Call ValidateRange(xCell)
Call ValidateRange(fCell)
Call ValidateRange(outCell)
xCellRE.Value = RangeAddress(xCell, "")
fCellRE.Value = RangeAddress(fCell, "")
OutcellRE.Value = RangeAddress(outCell, "")
LowTB.Value = a
UpTB.Value = b
etaTB.Value = Eta
epsTB.Value = Eps
minTB.Value = min
MaxTB.Value = max
End Sub
