# MSExcel-Numerical-Integration
This contains Visual Basic Code for numerical integration (using Romberg method) in Microsoft Excel. Enter *x* in any cell and enter the formula for *f(x)* in another cell. (You may use any number of intermediate cells to help compute *f(x)*).
In the form, provide the following data:
* the cell containing *x*
* the cell containing *f(x)*
* the cell in which to store the integral
* the lower and upper limits of integration, *a* and *b*

The VBA code computes the integral from *a* to *b* of *f(x) dx*

The code creates an auxilary function which stores its input argument into the *x* cell, reads the value contained in the *f(x)* cell and returns this value. The Romberg integration method is applied to this auxilary function. This allows the user to compute the integral of a function without explicitly defining the function in Visual Basic.
