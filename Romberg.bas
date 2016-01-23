Attribute VB_Name = "Romberg"
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
'This is part of the Quadrature form for numerical integration

'
'
Static Sub Quadrature()
Attribute Quadrature.VB_ProcData.VB_Invoke_Func = " \n14"
QuadratureForm.Show
End Sub
'
'
Static Function QuadratureDefaults(xCell, fCell, outCell, a, b, _
                Eps, Eta, min, max, store As Boolean)
'This is required only for the EstimDensity form
'It is defined here to allow the use of static variables
'whose values are remembered from invocation to invocation
'within the same Excel session
Static xCell0 As Range, fCell0 As Range, outCell0 As Range
Static a0, b0, eps0, eta0, min0, max0
If IsEmpty(eps0) Then eps0 = 0.001
If IsEmpty(eta0) Then eta0 = 0.001
If IsEmpty(min0) Then min0 = 3
If IsEmpty(max0) Then max0 = 10
'When called with store=True, the arguments are stored as new defaults
'Else the relevant global variables are set to their current defaults
If store Then
    Set xCell0 = xCell
    Set fCell0 = fCell
    Set outCell0 = outCell
    a0 = a
    b0 = b
    eps0 = Eps
    eta0 = Eta
    min0 = min
    max0 = max
Else
    Set xCell = xCell0
    Set fCell = fCell0
    Set outCell = outCell0
    a = a0
    b = b0
    Eps = eps0
    Eta = eta0
    min = min0
    max = max0
End If
End Function
