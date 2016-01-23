# MSExcel-Numerical-Integration
This contains Visual Basic Code for numerical integration (using Romberg method) in Microsoft Excel. Enter *x* in any cell and enter the formula for *f(x)* in another cell. (You may use any number of intermediate cells to help compute *f(x)*). The VBA code computes the integral from *a* to *b* of *f(x) dx*
The software consists of a form and some VBA code. The form allows the user to provide the following inputs:
* the cell containing *x*
* the cell containing *f(x)*
* the cell in which to store the integral
* the lower and upper limits of integration, *a* and *b*

The VBA code creates an auxilary function which stores its input argument into the *x* cell, reads the value contained in the *f(x)* cell and returns this as the function value. The Romberg integration method is applied to this auxilary function. This allows the user to compute the integral of a function without explicitly defining the function in Visual Basic.
