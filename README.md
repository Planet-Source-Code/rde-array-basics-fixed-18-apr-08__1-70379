<div align="center">

## Array Basics \(FIXED 18\-Apr\-08\)

<img src="PIC2008418311218955.jpg">
</div>

### Description

Array basics for beginners
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-array-basics-fixed-18-apr-08__1-70379/archive/master.zip)





### Source Code


<tt>
<hr width="95%" size="2" align="left">
<font color="#006600"><p nowrap>' There are two areas where arrays can be declared.</p>
<p nowrap>' Firstly in the General Declarations section. Arrays declared<br />
' here are seen by all procedures in this form or module:</p>
</font><p nowrap><font color="#000099">Private </font><font color="#660000">arrayName</font><font color="#330000">(</font><font color="#660000">[intLow To] intHigh</font><font color="#330000">) </font><font color="#000099">As </font><font color="#660000">dataType</p>
</font><font color="#006600"><p nowrap>' For example:</p>
</font><p nowrap><font color="#000099">Private </font><font color="#660000">lngArray</font><font color="#330000">(2 </font><font color="#000099">To </font><font color="#330000">4) </font><font color="#000099">As Long</p>
</font><font color="#006600"><p nowrap>' You can also declare the array as Public, making it visible to<br />
' all forms and modules in the application:</p>
</font><p nowrap><font color="#000099">Public </font><font color="#660000">arrayName</font><font color="#330000">(</font><font color="#660000">[intLow To] intHigh</font><font color="#330000">) </font><font color="#000099">As </font><font color="#660000">dataType</p>
</font><font color="#006600"><p nowrap>' For example:</p>
</font><p nowrap><font color="#000099">Public </font><font color="#660000">strArray</font><font color="#330000">(0 </font><font color="#000099">To </font><font color="#330000">200) </font><font color="#000099">As String</p>
</font><font color="#006600"><p nowrap>' Secondly, within procedures:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">arrayName</font><font color="#330000">(</font><font color="#660000">[intLow To] intHigh</font><font color="#330000">) </font><font color="#000099">As </font><font color="#660000">dataType</p>
</font><font color="#006600"><p nowrap>' For example:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">intArray</font><font color="#330000">(0 </font><font color="#000099">To </font><font color="#330000">9) </font><font color="#000099">As Integer</p>
</font><font color="#006600"><p nowrap>' You can also specify element size like a string variable:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">arrayName</font><font color="#330000">(</font><font color="#660000">[intLow To] intHigh</font><font color="#330000">) </font><font color="#000099">As </font><font color="#660000">dataType </font><font color="#330000">* </font><font color="#660000">intByteSize</p>
</font><font color="#006600"><p nowrap>' For example:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">strArray</font><font color="#330000">(3 </font><font color="#000099">To </font><font color="#330000">14) </font><font color="#000099">As String </font><font color="#330000">* 256</p>
</font><font color="#006600"><p nowrap>' Also, you can just specify the array length by omitting the<br />
' [intLow To] code shown above. This specifies the upper<br />
' element's index, not neccessarily the arrays length:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">ArrayName</font><font color="#330000">(</font><font color="#660000">intHigh</font><font color="#330000">) </font><font color="#000099">As </font><font color="#660000">dataType</p>
</font><font color="#006600"><p nowrap>' For example:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">intArray</font><font color="#330000">(9) </font><font color="#000099">As Integer</p>
</font><font color="#006600"><p nowrap>' This array will be indexed from 0 by default (so the array's<br />
' size will be one greater than intHigh) unless the following<br />
' line is added to the General Declarations section:</p>
</font><p nowrap><font color="#000099">Option Base </font><font color="#330000">1</p>
<hr width="95%" size="2" align="left">
</font><font color="#006600"><p nowrap>' You can also do this (for single dimensional arrays only):</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">scores </font><font color="#000099">As Variant<br />
</font><font color="#660000">scores </font><font color="#330000">= </font><font color="#660000">Array</font><font color="#330000">(81, 49, 80, 71, 92, 66)</p>
</font><font color="#006600"><p nowrap>' The above scores array will be indexed from 0 unless Option Base 1<br />
' is added to the General Declarations section.</p>
<hr width="95%" size="2" align="left">
<p nowrap>' VB supports static and dynamic arrays. Static arrays are fixed in<br />
' size and can't be changed at runtime, dynamic array sizes can.</p>
<p nowrap>' Static arrays are more memory efficient:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">array2</font><font color="#330000">(10 </font><font color="#000099">To </font><font color="#330000">20) </font><font color="#000099">As Integer</p>
</font><font color="#006600"><p nowrap>' Dynamic arrays do not have a size defined when initialized:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">array3</font><font color="#330000">() </font><font color="#000099">As Integer</p>
</font><font color="#006600"><p nowrap>' You must ReDim a dynamic array to change its size at runtime:</p>
</font><p nowrap><font color="#000099">ReDim </font><font color="#660000">array3</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">4) </font><font color="#000099">As Integer</p>
</font><font color="#006600"><p nowrap>' You can preserve the contents of the array elements when you<br />
' ReDim the array by using the Preserve keyword with ReDim:</p>
</font><p nowrap><font color="#000099">ReDim Preserve </font><font color="#660000">array3</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">9) </font><font color="#000099">As Integer</p>
</font><font color="#006600"><p nowrap>' You can also assign to a dynamic array directly from another<br />
' array without specifying size:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">array4</font><font color="#330000">() </font><font color="#000099">As Integer</p>
</font><p nowrap><font color="#660000">array4 </font><font color="#330000">= </font><font color="#660000">array3</p>
</font><font color="#006600"><p nowrap>' array4 will now be initialize with the size<br />
' (and element values if any) of array3.</p>
<hr width="95%" size="2" align="left">
<p nowrap>' Visual Basic allows you to use For Each ... Next to enumerate<br />
' the items in an array:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">element </font><font color="#000099">As Variant<br />
For Each </font><font color="#660000">element </font><font color="#000099">In </font><font color="#660000">array4</font><font color="#330000">()<br />
&#160; </font><font color="#006600">'code to process array elements sequentially<br />
</font><font color="#000099">Next</p>
</font><font color="#006600"><p nowrap>' Because arrays do not have a Count property you can use the<br />
' UBound and LBound methods to establish its length.</p>
<p nowrap>' You could use a loop like the following:</p>
</font><p nowrap><font color="#000099">For </font><font color="#660000">x </font><font color="#330000">= </font><font color="#000099">LBound</font><font color="#330000">(</font><font color="#660000">ArrayName</font><font color="#330000">()) </font><font color="#000099">To UBound</font><font color="#330000">(</font><font color="#660000">ArrayName</font><font color="#330000">())<br />
&#160; </font><font color="#006600">'code to process array items sequentially<br />
</font><font color="#000099">Next</p>
</font><font color="#006600"><p nowrap>' The arrays empty parentheses are not required, so the first<br />
' line of code above could be like this:</p>
</font><p nowrap><font color="#000099">For </font><font color="#660000">x </font><font color="#330000">= </font><font color="#000099">LBound</font><font color="#330000">(</font><font color="#660000">ArrayName</font><font color="#330000">) </font><font color="#000099">To UBound</font><font color="#330000">(</font><font color="#660000">ArrayName</font><font color="#330000">)</p>
</font><font color="#006600"><p nowrap>' To establish the length you could also use the following code<br />
' that subtracts the LowerBound value from the UpperBound value,<br />
' then adds one to the result because the array elements include<br />
' both upper and lower, then returns the result As Integer:</p>
</font><p nowrap><font color="#000099">Function </font><font color="#660000">GetCount</font><font color="#330000">(</font><font color="#660000">AnyArray </font><font color="#000099">As Variant</font><font color="#330000">) </font><font color="#000099">As Integer<br />
&#160; On </font><font color="#660000">Error </font><font color="#000099">Resume Next<br />
&#160; Dim </font><font color="#660000">length </font><font color="#000099">As Integer<br />
&#160; </font><font color="#660000">length </font><font color="#330000">= </font><font color="#000099">UBound</font><font color="#330000">(</font><font color="#660000">AnyArray</font><font color="#330000">) - </font><font color="#000099">LBound</font><font color="#330000">(</font><font color="#660000">AnyArray</font><font color="#330000">) + 1<br />
&#160; </font><font color="#660000">GetCount </font><font color="#330000">= </font><font color="#660000">length             </font><font color="#006600">' + 1 = inclusive<br />
</font><font color="#000099">End Function</p>
</font><font color="#006600"><p nowrap>' If you know the data type of the array you could declare<br />
' the AnyArray argument as that type instead of As Variant<br />
' to improve performance:</p>
</font><p nowrap><font color="#000099">Function </font><font color="#660000">GetCount</font><font color="#330000">(</font><font color="#660000">sngArray() </font><font color="#000099">As Single</font><font color="#330000">) </font><font color="#000099">As Integer<br />
&#160; On </font><font color="#660000">Error </font><font color="#000099">Resume Next<br />
&#160; Dim </font><font color="#660000">length </font><font color="#000099">As Integer<br />
&#160; </font><font color="#660000">length </font><font color="#330000">= </font><font color="#000099">UBound</font><font color="#330000">(</font><font color="#660000">sngArray</font><font color="#330000">) - </font><font color="#000099">LBound</font><font color="#330000">(</font><font color="#660000">sngArray</font><font color="#330000">) + 1<br />
&#160; </font><font color="#660000">GetCount </font><font color="#330000">= </font><font color="#660000">length             </font><font color="#006600">' + 1 = inclusive<br />
</font><font color="#000099">End Function</p>
</font><font color="#006600"><p nowrap>' Calling the function is as easy as:</p>
</font><p nowrap><font color="#660000">intCount </font><font color="#330000">= </font><font color="#660000">GetCount</font><font color="#330000">(</font><font color="#660000">myArray</font><font color="#330000">())</p>
<hr width="95%" size="2" align="left">
</font><font color="#006600"><p nowrap><p nowrap>' Multi-dimensional arrays</p>
<p nowrap>' Each dimension of the array must contain the same data type.</p>
<p nowrap>' The second-last and last dimensions of a multi-dimensional array<br />
' are normally considered to be a Row and a Column respectively.</p>
</font><p nowrap><font color="#000099">Private </font><font color="#660000">multiArray</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">5, 1 </font><font color="#000099">To </font><font color="#330000">3) </font><font color="#000099">As Integer</p>
</font><font color="#006600"><p nowrap>' So for a two dimensional array in VB the dimension (row) is<br />
' defined first, and the number of elements (cols) for each<br />
' dimension defined second:</p>
</font><p nowrap><font color="#000099">Public </font><font color="#660000">2DArray</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">2, 1 </font><font color="#000099">To </font><font color="#330000">8) </font><font color="#000099">As Long</p>
</font><font color="#006600"><p nowrap>' This array is a two dimensional array containing 8 elements<br />
' in each dimension:</p>
</font><p nowrap><font color="#000099">Dim </font><font color="#660000">i1 </font><font color="#000099">As Integer</font><font color="#330000">, </font><font color="#660000">i2 </font><font color="#000099">As Integer<br />
For </font><font color="#660000">i1 </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">2<br />
&#160; </font><font color="#000099">For </font><font color="#660000">i2 </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">8<br />
&#160; &#160; </font><font color="#660000">2DArray</font><font color="#330000">(</font><font color="#660000">i1</font><font color="#330000">, </font><font color="#660000">i2</font><font color="#330000">) = </font><font color="#666666">&quot;Cell &quot; </font><font color="#330000">&amp; </font><font color="#000099">CStr</font><font color="#330000">(</font><font color="#660000">i1</font><font color="#330000">) &amp; </font><font color="#666666">&quot;,&quot; </font><font color="#330000">&amp; </font><font color="#000099">CStr</font><font color="#330000">(</font><font color="#660000">i2</font><font color="#330000">)<br />
&#160; </font><font color="#000099">Next<br />
Next</p>
</font><font color="#006600"><p nowrap>' You can also assign arrays to the elements of other arrays to<br />
' create multi-dimensional arrays:</p>
</font><p nowrap><font color="#000099">Public Sub </font><font color="#660000">CreateMultiArray</font><font color="#330000">()<br />
&#160; </font><font color="#000099">Dim </font><font color="#660000">intX </font><font color="#000099">As Integer </font><font color="#006600">' Declare counter variable</p>
<p nowrap>&#160; ' Declare and populate an integer array<br />
&#160; </font><font color="#000099">Dim </font><font color="#660000">countersA</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">4) </font><font color="#000099">As Integer</p>
<p nowrap>&#160; For </font><font color="#660000">intX </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">4<br />
&#160; &#160; </font><font color="#660000">countersA</font><font color="#330000">(</font><font color="#660000">intX</font><font color="#330000">) = </font><font color="#660000">intX<br />
&#160; </font><font color="#000099">Next </font><font color="#660000">intX</font></p>
<p nowrap>&#160; <font color="#006600">' Declare and populate a string array</font><br />
&#160; <font color="#000099">Dim </font><font color="#660000">countersB</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">4) </font><font color="#000099">As String</p>
<p nowrap>&#160; For </font><font color="#660000">intX </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">4<br />
&#160; &#160; </font><font color="#660000">countersB</font><font color="#330000">(</font><font color="#660000">intX</font><font color="#330000">) = </font><font color="#666666">&quot;hello&quot;<br />
&#160; </font><font color="#000099">Next </font><font color="#660000">intX</p>
<p nowrap>&#160; </font><font color="#006600">' Declare a new two-member array</font><br />
&#160; <font color="#000099">Dim </font><font color="#660000">arrX</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">2) </font><font color="#000099">As Variant</font></p>
<p nowrap>&#160; <font color="#006600">' Populate the array with other arrays</font><br />
&#160; <font color="#660000">arrX</font><font color="#330000">(1) = </font><font color="#660000">countersA</font><font color="#330000">()</font><br />
&#160; <font color="#660000">arrX</font><font color="#330000">(2) = </font><font color="#660000">countersB</font><font color="#330000">()</font></p>
<p nowrap>&#160; <font color="#006600">' Display a member of each array</font><br />
&#160; <font color="#660000">MsgBox arrX</font><font color="#330000">(1)(2)</font><br />
&#160; <font color="#660000">MsgBox arrX</font><font color="#330000">(2)(3)</font></p>
<font color="#000099"><p nowrap>End Sub</p></font>
<font color="#006600"><p nowrap>' To increase the size of an array without losing its current<br />
' values use the Preserve keyword:</p></font>
<p nowrap><font color="#000099">ReDim Preserve </font><font color="#660000">DynArray</font><font color="#330000">(</font><font color="#000099">UBound</font><font color="#330000">(</font><font color="#660000">DynArray</font><font color="#330000">) + 1)</font></p>
<p nowrap><font color="#006600">' Only the upper bound of the last dimension in a multi-dimensional<br />
' array can be changed when you use the Preserve keyword; if you<br />
' change any of the other dimensions, or the lower bound of the<br />
' last dimension, a run-time error occurs.</p>
<p nowrap>' Thus, you can use code like this:</p></font>
<p nowrap><font color="#000099">ReDim Preserve </font><font color="#660000">Matrix</font><font color="#330000">(10, </font><font color="#000099">UBound</font><font color="#330000">(</font><font color="#660000">Matrix</font><font color="#330000">, 2) + 1)</font></p>
<p nowrap><font color="#006600">' But you cannot use this code:</font></p>
<p nowrap><font color="#000099">ReDim Preserve </font><font color="#660000">Matrix</font><font color="#330000">(</font><font color="#000099">UBound</font><font color="#330000">(</font><font color="#660000">Matrix</font><font color="#330000">, 1) + 1, 10)</font></p>
<hr width="95%" size="2" align="left">
<font color="#006600"><p nowrap>' Multi-dimensional array demonstration</p>
<p nowrap>' The following code can be copied and pasted into the form of<br />
' a new project and after creating three command buttons named<br />
' cmdFill, cmdShow and cmdMulti you can run the program:</p>
</font><font color="#000099"><p nowrap>Option Explicit</p>
<p nowrap>Option Base </font><font color="#330000">1</font></p>
<p nowrap><font color="#000099">Private </font><font color="#660000">multiArray</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">5, 1 </font><font color="#000099">To </font><font color="#330000">3) </font><font color="#000099">As String<br />
Private </font><font color="#660000">counter </font><font color="#000099">As Integer</p>
<p nowrap>Private Sub </font><font color="#660000">cmdFill_Click</font><font color="#330000">()</font></p>
<p nowrap>&#160; <font color="#000099">Dim </font><font color="#660000">idx1 </font><font color="#000099">As Integer</font><font color="#330000">, </font><font color="#660000">idx2 </font><font color="#000099">As Integer</p>
<p nowrap>&#160; For </font><font color="#660000">idx1 </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">5<br />
&#160; &#160; </font><font color="#000099">For </font><font color="#660000">idx2 </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">3<br />
&#160; &#160; &#160; </font><font color="#660000">counter </font><font color="#330000">= </font><font color="#660000">counter </font><font color="#330000">+ 1<br />
&#160; &#160; &#160; </font><font color="#660000">multiArray</font><font color="#330000">(</font><font color="#660000">idx1</font><font color="#330000">, </font><font color="#660000">idx2</font><font color="#330000">) = </font><font color="#666666">&quot;Cell &quot; </font><font color="#330000">&amp; </font><font color="#000099">CStr</font><font color="#330000">(</font><font color="#660000">counter</font><font color="#330000">)<br />
&#160; &#160; </font><font color="#000099">Next<br />
&#160; Next</p>
<p nowrap>End Sub</p>
<p nowrap>Private Sub </font><font color="#660000">cmdShow_Click</font><font color="#330000">()</font></p>
<p nowrap>&#160; <font color="#660000">Me</font><font color="#330000">.</font><font color="#660000">Refresh<br />
&#160; CurrentY </font><font color="#330000">= 100</font><br />
&#160; <font color="#000099">Print </font><font color="#666666">&quot; Dim multiarray(1 To 5, 1 To 3) As String&quot;</font><br />
&#160; <font color="#660000">CurrentY </font><font color="#330000">= 400</font></p>
<p nowrap>&#160; <font color="#000099">Dim </font><font color="#660000">idx1 </font><font color="#000099">As Integer</font><font color="#330000">, </font><font color="#660000">idx2 </font><font color="#000099">As Integer</p>
<p nowrap>&#160; For </font><font color="#660000">idx1 </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">5</font><br />
&#160; &#160; <font color="#000099">For </font><font color="#660000">idx2 </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">3</font><br />
&#160; &#160; &#160; <font color="#000099">Print </font><font color="#666666">&quot;  multiarray(&quot; </font><font color="#330000">&amp; </font><font color="#660000">_</font><br />
&#160; &#160; &#160; &#160; <font color="#000099">CStr</font><font color="#330000">(</font><font color="#660000">idx1</font><font color="#330000">) &amp; </font><font color="#666666">&quot;,&quot; </font><font color="#330000">&amp; </font><font color="#000099">CStr</font><font color="#330000">(</font><font color="#660000">idx2</font><font color="#330000">) &amp; </font><font color="#666666">&quot;) = &quot; </font><font color="#330000">&amp; </font><font color="#660000">_<br />
&#160; &#160; &#160; &#160; Chr</font><font color="#330000">(34) &amp; </font><font color="#660000">multiArray</font><font color="#330000">(</font><font color="#660000">idx1</font><font color="#330000">, </font><font color="#660000">idx2</font><font color="#330000">) &amp; </font><font color="#660000">Chr</font><font color="#330000">(34)</font><br />
&#160; &#160; <font color="#000099">Next<br />
&#160; Next</p>
<p nowrap>End Sub</p>
<p nowrap>Private Sub </font><font color="#660000">cmdMulti_Click</font><font color="#330000">()</font><br />
&#160; <font color="#000099">Dim </font><font color="#660000">intX </font><font color="#000099">As Integer </font><font color="#006600">' Declare counter variable</p>
<p nowrap>&#160; ' Declare and populate an integer array</font><br />
&#160; <font color="#000099">Dim </font><font color="#660000">countersA</font><font color="#330000">(4) </font><font color="#000099">As Integer</p>
<p nowrap>&#160; For </font><font color="#660000">intX </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">4</font><br />
&#160; &#160; <font color="#660000">countersA</font><font color="#330000">(</font><font color="#660000">intX</font><font color="#330000">) = </font><font color="#660000">intX</font><br />
&#160; <font color="#000099">Next </font><font color="#660000">intX</font></p>
<p nowrap>&#160; <font color="#006600">' Declare and populate a string array</font><br />
&#160; <font color="#000099">Dim </font><font color="#660000">countersB</font><font color="#330000">(4) </font><font color="#000099">As String</p>
<p nowrap>&#160; For </font><font color="#660000">intX </font><font color="#330000">= 1 </font><font color="#000099">To </font><font color="#330000">4</font><br />
&#160; &#160; <font color="#660000">countersB</font><font color="#330000">(</font><font color="#660000">intX</font><font color="#330000">) = </font><font color="#666666">&quot;hello&quot;</font><br />
&#160; <font color="#000099">Next </font><font color="#660000">intX</font></p>
<p nowrap>&#160; <font color="#006600">' Declare a new two-member array</font><br />
&#160; <font color="#000099">Dim </font><font color="#660000">arrX</font><font color="#330000">(1 </font><font color="#000099">To </font><font color="#330000">2) </font><font color="#000099">As Variant</font></p>
<p nowrap>&#160; <font color="#006600">' Populate the array with other arrays</font><br />
&#160; <font color="#660000">arrX</font><font color="#330000">(1) = </font><font color="#660000">countersA</font><font color="#330000">()</font><br />
&#160; <font color="#660000">arrX</font><font color="#330000">(2) = </font><font color="#660000">countersB</font><font color="#330000">()</font></p>
<p nowrap>&#160; <font color="#006600">' Display a member of each array</font><br />
&#160; <font color="#660000">MsgBox arrX</font><font color="#330000">(1)(2)</font><br />
&#160; <font color="#660000">MsgBox arrX</font><font color="#330000">(2)(3)</font><br />
<font color="#000099">End Sub</font></p>
<hr width="95%" size="2" align="left"></tt>

