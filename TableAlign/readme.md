# Step 1

Here is the [code](TableAlignMacro.txt)

# Step 2

**Understand the code**

<code> Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed) </code> <br>
<code> Selection.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPoints </code> <br>
<code> Selection.Tables(1).Columns(1).PreferredWidth = CentimetersToPoints(10) </code> <br>
<code> Selection.Tables(1).Columns(2).PreferredWidth = CentimetersToPoints(2.5) </code> <br>
<code> Selection.Tables(1).Columns(3).PreferredWidth = CentimetersToPoints(0.5) </code> <br>
<code> Selection.Tables(1).Columns(4).PreferredWidth = CentimetersToPoints(2.5) </code> <br>

**According to the code;**

- The table should have four columns.
- Once macro is clicked, the width of;
	- the first column will be set to 10cm, 
	- the second will be set to 2.5cm and so on...
- So the table will look like this (with sample data)
	- ![](x.%20assets/fig1.png)

# Examples

**Let's take a look at a bigger table**

![](x.%20assets/fig2.png)

Clicking the macro with the code mentioned above will give an <font color=red><b>error</b></font>, because this table has 6 columns as opposed to the 4 columns the macro will work for.

So here, the macro is to be adjusted.

Make a new macro. Copy the code mentioned above and make similar changes;

- Add extra lines of code for the extra columns
- And change the width of the columns for these extra lines of code accordingly

Code to fix the table in this example:

<code> Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed) </code> <br>
<code> Selection.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPoints </code> <br>
<code> Selection.Tables(1).Columns(1).PreferredWidth = CentimetersToPoints(9) </code> <br>
<code> Selection.Tables(1).Columns(2).PreferredWidth = CentimetersToPoints(1.75) </code> <br>
<code> Selection.Tables(1).Columns(3).PreferredWidth = CentimetersToPoints(0.5) </code> <br>
<code> Selection.Tables(1).Columns(4).PreferredWidth = CentimetersToPoints(2.5) </code> <br>
<code> Selection.Tables(1).Columns(5).PreferredWidth = CentimetersToPoints(0.5) </code> <br>
<code> Selection.Tables(1).Columns(6).PreferredWidth = CentimetersToPoints(2.5) </code>

and this will fix the table to give:

![](x.%20assets/fig3.png)

