VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'hasha ,hashm, hashurl 三个部分组成, ComputeHash_2 在win8.1 with update x64, office 2016 x86出现内存溢出错误, win8.1 x, office x64, win7 sp1 x86, office2016 x32测试通过, 文件大小344M
'property的设置,https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/property-set-statement
'--------------https://docs.microsoft.com/zh-cn/office/vba/language/concepts/getting-started/writing-a-property-procedure
'Property Set 可用于对象，而 Property Let 不可以
'Property Get    属性 Getpropname(1, ..., n)为类型
'Property Let    Property Letpropname(1, ...,,,, n, n + 1)
'Property Set    属性集propname(1, ..., n, n + 1)
'Property Get 过程声明较相关的 Property Let 和 Property Set 声明少一个参数。 属性 Get过程的数据类型必须与相关的Property Let和Property Set声明中的最后一个参数 (n + 1) 的数据类型相同。
'例如，如果您声明以下 Property Let 过程，则 Property Get 声明必须使用与 Property Let 过程中的参数具有相同的名称和数据类型的参数。
'如果未使用**public、 Private 或Friend** 显式指定, 则默认情况下属性过程是公共的。 如果未使用**Static** , 则不会在两个调用之间保留局部变量的值。
'Friend 关键字只能用于类模块。 但是，项目的任何模块中的过程均可以访问 Friend 过程。 Friend 过程既不会显示在其父类的类型库中，也不会成为稍后可绑定的 Friend 过程。
'所有可执行代码必须位于过程中。 不能在另一个property、 Sub 或**Function** 过程中定义property Set过程。
'Exit Property 语句导致立即退出属性集过程。 程序将继续执行已调用 Property Set 过程的语句后面的语句。 任意数量的 Exit Property 语句可以出现在 Property Set 过程的任意位置。
'与 Function 和 Property Get 过程一样， Property Set 过程是一个单独的过程，它可以获取参数、执行一系列语句并更改其参数的值。
'但是, 与函数和Property Get过程不同, 这两个函数都返回一个值, 您只能在对象引用分配或**Set** 语句的左侧使用Property Set过程。
'--------------------------https://stackoverflow.com/questions/9670022/what-is-the-difference-between-public-property-friend-and-public-variable-in-vb
Private Const CryptographyUrlx As String = "https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography?view=netframework-4.8"
Private AHash As New HashA '一定要加new, 否则会出现对象错误
Private UHash As New HashUrl

Friend Property Get Algorithmx() As HashA
    Set Algorithmx = AHash
End Property
'Property Set 声明中的最后一个参数的数据类型必须为对象类型或 Variant。
Friend Property Set Algorithmx(ByVal obj As HashA)
    Set AHash = obj
End Property

Friend Property Get AlgorithmxUrl() As HashUrl
    Set AlgorithmxUrl = UHash
End Property

Friend Property Set AlgorithmxUrl(ByVal obj As HashUrl)
    Set UHash = obj
End Property

Friend Property Get CryptographyUrl() As String
    CryptographyUrl = CryptographyUrlx
End Property

Private Sub Class_Terminate()
    Set AHash = Nothing
    Set UHash = Nothing
End Sub
