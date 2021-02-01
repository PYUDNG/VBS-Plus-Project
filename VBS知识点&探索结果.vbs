有关引用
	[VariantName]不受变量命名规则的限制，可以含有特殊字符或者直接命名为VBS关键字或常量形式名称，但是VariantName仍然不能含有反方括号"]"和换行符（也许其他的奇葩字符也不可以，比如Chr(8)？待测试。）
	[VariantName]和VariantName引用的是同一个变量
	Function和Sub可以重定义，Class不可以
	ReDim语句可以覆盖原先变量的值和类型
	Dim语句定义在ExecuteGlobal()里面重定义变量，但这不会改变已有变量的值和类型
	Erase语句清除数组变量后，该变量会成为一个空数组，最大下标为0，IsArray()返回True，VarType()返回8204

有关压缩的流程控制语句
	If语句块可以压缩在一个逻辑行里，Select Case中的Case块也可以压缩在一个逻辑行里
	压缩的If语句中，Statements之间必须有冒号划分内部逻辑行，而Statement与流程判断语句本身之间无需冒号（有无冒号都行）
	压缩的If语句中，End If是非必要的，有没有都行，但如果存在If嵌套，那么非最外层的End If就是必须的了（压缩的最外层If语句可以没有End If）
	压缩的If语句可以通过物理行连接符"_"写到多个逻辑行里
	VBS本身不支持在压缩的If语句内部的ElseIf语句块
	Do Loop、For Next、While Wend不可以无冒号压缩，有冒号压缩也必须完整（Do必须有与之对应的Loop，For必须有与之对应的Next，While必须有与之对应的Wend）

有关函数和子程序
	Function F()：
		Call F(Args...)：传值传址根据定义中的ByVal、ByRef关键字确定。没有关键字，默认ByRef。
		Var = F(Args...)：同Call F(Args...)。
		F(Args...)：无论定义是ByVal还是ByRef抑或是没有关键字，都是ByRef。
	Sub S()：
		Call S(Args...)：同Call F(Args...)。
		Var = S(Args...)：同Call F(Args...)。
		S(Args...)：同F(Args...)。

有关逻辑行与物理行
	":"分割物理行为多个逻辑行，"_"把多个物理行连接成一个逻辑行
	":"和"_"可以混用，比如说把3个逻辑行写到4个物理行里，使其中任何一个物理行都不包含一个完整的逻辑行（谁会这么写代码啊？支持这种写法真是累死我了）
	当"_"在行尾时：
		如果和前面明显是分开的（比如，它前面一个字符是空格" "，英文逗号","，或者加号"+"等），就把他看作物理行连接符；
		如果不是和前面明显分开，就把他看作前一个单词的一部分，不管这个单词作为名称是否真的存在（比如If xxx Or_的最后一个字符是"_"，但是解释器会把"Or_"当成一个整体）
		