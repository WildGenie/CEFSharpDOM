Imports CefSharp
Imports CefSharp.WinForms

Public Class CEFBrowser
	Inherits ChromiumWebBrowser
	Implements IDownloadHandler

	Public Event DocumentComplete()
	Public Event DocumentChanged()
	Public Event DOMEvent(EventName As String, SourceIndex As Integer)

	Public Document As New CEFDocument(Me)
	Private ScriptDeferred As Boolean
	Private DeferredScript As New List(Of String)

	'Receives messages setup with addEventListener
	Private WithEvents CEFMessage As New CEFMessage

	'Add any listeners from here (comma separated list)
	Public Sub AddListeners(EventNames As String)
		Dim Code = ""
		For Each Item In EventNames.Split(",")
			Code &= String.Format("document.addEventListener('{0}',function(e){{_CmX.msg('{0}',e.target.srcIndex())}},true);", Item)
		Next
		Me.ExecScript(Code)
	End Sub

	Private Sub CEFBrowser_IsBrowserInitializedChanged(sender As Object, e As IsBrowserInitializedChangedEventArgs) Handles Me.IsBrowserInitializedChanged
		If e.IsBrowserInitialized Then
			Me.JavascriptObjectRepository.Register("_CmX", CEFMessage, True)
		End If
	End Sub

	'Raise listener eventwith sourceIndex of activeElement
	Private Sub CEFMessage_DOMEvent(Msg As String, Idx As Integer) Handles CEFMessage.DOMEvent
		BeginInvoke(Sub() RaiseDOMEvent(Msg, Idx))
	End Sub

	Private Sub ChromeBrowser_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
		TChange.Stop()
	End Sub

	Private Sub ChromeBrowser_FrameLoadEnd(sender As Object, e As FrameLoadEndEventArgs) Handles Me.FrameLoadEnd
		If e.Frame.IsMain Then
			DeferScript(True)
			ExecScript("CefSharp.BindObjectAsync('_CmX')")
			ExecScript("var _DsX;HTMLElement.prototype.srcIndex=function(){return _DsX.indexOf(this)}")
			ExecScript("var cf={childList:true};var cb=function(l){_CmX.msg('change',0)};var ob=new MutationObserver(cb);ob.observe(document.body,cf)")
			DeferScript(False)
			ContentChanged(True)
		End If
	End Sub

	Private Async Sub ContentChanged(Main As Boolean)
		Dim Res = Await Me.EvaluateScriptAsync("(function(){_DsX=Array.from(document.all);var r='';for(i=0;i<_DsX.length;i++){r+=_DsX[i].tagName+' '} return r})()")
		If Res.Success Then
			Invoke(Sub() Document.Load(Res.Result))
			Invoke(Sub() RaiseChangeEvent(Main))
			Debug.Print("reloaded")
		End If
	End Sub

	'Defer scripts
	Public Sub DeferScript(Defer As Boolean)
		If Defer Then
			ScriptDeferred = True
			DeferredScript.Clear()
		Else
			ScriptDeferred = False
			ExecScript(Join(DeferredScript.ToArray, ";"))
			DeferredScript.Clear()
		End If
	End Sub

	'Evalscript: wraps EvalScriptAsync to catch errors and save having to escape {}
	Public Function EvalScript(Code As String, ParamArray Params() As String) As String
		For i = 0 To Params.UBound
			Code = Code.Replace("{" & i & "}", Params(i))
		Next
		Return EvalScriptSync("(function(){try{" & Code & "}catch(e){console.log(e.message);return ''}})()")
	End Function

	'Wait for EvalScriptAsync
	Private Function EvalScriptSync(Code As String) As String
		If Me.IsBrowserInitialized Then
			Using Ret = Me.EvaluateScriptAsync(Code)
				'Not using wait, so doesn't lock on breakpoints
				Do
					Application.DoEvents()
				Loop Until Ret.IsCompleted
				Return Ret.Result.Result.ToString.Trim
			End Using
		Else
			Return ""
		End If
	End Function

	Friend Sub ExecScript(Code As String)
		If Me.IsBrowserInitialized Then
			If ScriptDeferred Then
				DeferredScript.Add(Code)
			Else
				ExecuteScriptAsync(Code)
			End If
		End If
	End Sub

	Private Sub OnBeforeDownload(browser As IBrowser, downloadItem As DownloadItem, callback As IBeforeDownloadCallback) Implements IDownloadHandler.OnBeforeDownload
	End Sub

	Private Sub OnDownloadUpdated(browser As IBrowser, downloadItem As DownloadItem, callback As IDownloadItemCallback) Implements IDownloadHandler.OnDownloadUpdated
	End Sub

	'Document has loaded or changed
	Private Sub RaiseChangeEvent(Main As Boolean)
		On Error Resume Next
		If Main Then
			RaiseEvent DocumentComplete()
		Else
			RaiseEvent DocumentChanged()
		End If
	End Sub

	'Raise DOM event added with addEventListener
	Private Sub RaiseDOMEvent(Msg As String, Idx As String)
		If Msg = "change" Then
			TChange.Stop()
			TChange.Start()
		Else
			RaiseEvent DOMEvent(Msg, Idx)
		End If
	End Sub

	'Times out after 150ms to prevent MutationObserver events is quick succession
	Private Sub TChange_Tick(sender As Object, e As EventArgs) Handles TChange.Tick
		TChange.Stop()
		ContentChanged(False)
	End Sub

End Class

'Wraps basic document functionality in a class
Public Class CEFDocument
	Public Event Message(Message As String, Element As CEFElement)
	Public ReadOnly Chrome As CEFBrowser
	Public ReadOnly All As New List(Of CEFElement)

	Public ReadOnly Property ActiveElement() As CEFElement
		Get
			Try
				Return All(Chrome.EvalScript("return document.activeElement.srcIndex()"))
			Catch ex As Exception
				Return Nothing
			End Try
		End Get
	End Property

	Public ReadOnly Property Body() As CEFElement
		Get
			For Each Item In Me.All
				If Item.TagName = "body" Then
					Return All(Item.SourceIndex)
				End If
			Next
			Return Nothing
		End Get
	End Property

	Public ReadOnly Property ElementById(ID As String) As CEFElement
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.getElementById('{0}')", ID)
			Dim Ret = CInt(Me.Chrome.EvalScript(Code))
			Return Me.All(Ret)
		End Get
	End Property

	Public ReadOnly Property ElementsByTagName(Name As String) As List(Of CEFElement)
		Get
			On Error Resume Next
			ElementsByTagName = New List(Of CEFElement)
			Dim Code = "var a=document.getElementsByTagName('{0}');var r='';for(i=0;i<a.length;i++){r+=a[i].srcIndex()+' '} return r"
			Dim Ret = CStr(Me.Chrome.EvalScript(Code, Name))
			If Ret.Length Then
				For Each Item In Ret.Split
					ElementsByTagName.Add(Me.All(Item))
				Next
			End If
		End Get
	End Property

	Friend Sub Load(List As String)
		'Load basic DOM tree details into our list
		All.Clear()
		Dim i As Integer
		For Each Item In List.ToLower.Split
			All.Add(New CEFElement(Me, Item, i))
			i += 1
		Next
	End Sub

	Friend Sub New(CEFBrowser As CEFBrowser)
		Chrome = CEFBrowser
	End Sub

	Public ReadOnly Property QuerySelector(Query As String) As CEFElement
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.querySelector(""{0}"").srcIndex()", Query)
			Dim Ret = CInt(Me.Chrome.EvalScript(Code))
			Return Me.All(Ret)
		End Get
	End Property

	Public ReadOnly Property QuerySelectorAll(Query As String) As List(Of CEFElement)
		Get
			On Error Resume Next
			QuerySelectorAll = New List(Of CEFElement)
			Dim Code = "var a=document.querySelectorAll(""{1}"");var r='';for(i=0;i<a.length;i++){r+=a[i].srcIndex()+' '} return r"
			Dim Ret = CStr(Me.Chrome.EvalScript(Code, Query))
			If Ret.Length Then
				For Each Item In Ret.Split
					QuerySelectorAll.Add(Me.All(Item))
				Next
			End If
		End Get
	End Property

End Class

'Wraps basic element functionality in a class
Public Class CEFElement
	Public ReadOnly Document As CEFDocument
	Public ReadOnly SourceIndex As Integer
	Public ReadOnly TagName As String

	Public Sub [Invoke](Name As String)
		On Error Resume Next
		Dim Code = String.Format("document.all[{0}].{1}()", SourceIndex, Name)
		Me.Document.Chrome.ExecScript(Code)
	End Sub

	Default Public Property [Property](Name As String) As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.all[{0}].{1}", SourceIndex, Name)
			Return Me.Document.Chrome.EvalScript(Code)
		End Get
		Set(value As String)
			On Error Resume Next
			Dim Code = String.Format("document.all['{0}'].{1}='{2}'", SourceIndex, Name, value)
			Me.Document.Chrome.ExecScript(Code)
		End Set
	End Property

	Public ReadOnly Property Ancestor(TagName As String) As CEFElement
		Get
			On Error Resume Next
			Dim Code = "for(i={0};i>0;i--){if(document.all[i].tagName=='{1}'){return i}}"
			Dim Ret = CInt(Me.Document.Chrome.EvalScript(Code, SourceIndex, TagName.ToUpper))
			Return Document.All(Ret)
		End Get
	End Property

	Public Property Attribute(Name As String) As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.all[{0}].getAttribute('{1}')", SourceIndex, Name)
			Dim Json = Me.Document.Chrome.EvalScript(Code)
			Return ""
		End Get
		Set(value As String)
			On Error Resume Next
			Dim Code = String.Format("document.all[{0}].setAttribute('{1}','{2}')", SourceIndex, Name, value)
			Me.Document.Chrome.ExecScript(Code)
		End Set
	End Property

	Public ReadOnly Property BoundingClientRect() As Rectangle
		Get
			On Error Resume Next
			Dim Code = "var c=document.all[{0}].getBoundingClientRect();return c.left+' '+c.top+' '+c.width+' '+c.height"
			Dim Ret = CStr(Me.Document.Chrome.EvalScript(Code, SourceIndex)).Split
			If Ret.Length = 4 Then
				Return New Rectangle(Ret(0), Ret(1), Ret(2), Ret(3))
			End If
		End Get
	End Property

	Public ReadOnly Property Children() As List(Of CEFElement)
		Get
			On Error Resume Next
			Children = New List(Of CEFElement)
			Dim Code = "var a=document.all[{0}].children;var r='';for(i=0;i<a.length;i++){r+=a[i].srcIndex()+' '} return r"
			Dim Ret = CStr(Me.Document.Chrome.EvalScript(Code, SourceIndex))
			If Ret.Length Then
				For Each Item In Ret.Split
					Children.Add(Me.Document.All(Item))
				Next
			End If
		End Get
	End Property

	Public ReadOnly Property ComputedStyle() As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return window.getComputedStyle(document.all[{0}],null).cssText", SourceIndex)
			Return Me.Document.Chrome.EvalScript(Code)
		End Get
	End Property

	Public ReadOnly Property ComputedStyle(Style As String) As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return window.getComputedStyle(document.all[{0}],null).getPropertyValue('{1}')", SourceIndex, Style)
			Return Me.Document.Chrome.EvalScript(Code)
		End Get
	End Property

	Public ReadOnly Property Descendants(Name As String) As List(Of CEFElement)
		Get
			On Error Resume Next
			Descendants = New List(Of CEFElement)
			Dim Code = "var a=document.all[{0}].getElementsByTagName('{1}');var r='';for(i=0;i<a.length;i++){r+=a[i].srcIndex()+' '} return r"
			Dim Ret = CStr(Me.Document.Chrome.EvalScript(Code, SourceIndex, Name))
			If Ret.Length Then
				For Each Item In Ret.Split
					Descendants.Add(Me.Document.All(Item))
				Next
			End If
		End Get
	End Property

	Public Property InnerHTML() As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.all[{0}].innerHTML", SourceIndex)
			Return Me.Document.Chrome.EvalScript(Code)
		End Get
		Set(value As String)
			On Error Resume Next
			Dim Code = String.Format("document.all[{0}].innerHTML=""{1}""", SourceIndex, value)
			Me.Document.Chrome.ExecScript(Code)
		End Set
	End Property

	Public Property InnerText() As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.all[{0}].innerText", SourceIndex)
			Return Me.Document.Chrome.EvalScript(Code)
		End Get
		Set(value As String)
			On Error Resume Next
			Dim Code = String.Format("document.all[{0}].innerText=""{1}""", SourceIndex, value)
			Me.Document.Chrome.ExecScript(Code)
		End Set
	End Property

	Public Sub InsertAdjacentHTML(Where As String, HTML As String)
		On Error Resume Next
		Dim Code = String.Format("document.all[{0}].insertAdjacentHTML(""{1}"",""{2}"")", SourceIndex, Where, HTML)
		Me.Document.Chrome.ExecScript(Code)
	End Sub

	Public Sub New(CEFDocument As CEFDocument, Name As String, Index As Integer)
		Document = CEFDocument
		SourceIndex = Index
		TagName = Name
	End Sub

	Public Property OuterHTML() As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.all[{0}].outerHTML", SourceIndex)
			Return Me.Document.Chrome.EvalScript(Code)
		End Get
		Set(value As String)
			On Error Resume Next
			Dim Code = String.Format("document.all[{0}].outerHTML=""{1}""", SourceIndex, value)
			Me.Document.Chrome.ExecScript(Code)
		End Set
	End Property

	Public ReadOnly Property Parent(Optional Count As Integer = 1) As CEFElement
		Get
			On Error Resume Next
			Dim Code = "var e=document.all[{0}];for(i=0;i<{1};i++){e=e.parentElement} return e.srcIndex()"
			Dim Ret = CInt(Me.Document.Chrome.EvalScript(Code, SourceIndex, Count))
			Return Document.All(Ret)
		End Get
	End Property

	Public ReadOnly Property QuerySelector(Query As String) As CEFElement
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.all[{0}].querySelector(""{1}"").srcIndex()", SourceIndex, Query)
			Dim Ret = CInt(Me.Document.Chrome.EvalScript(Code))
			Return Me.Document.All(Ret)
		End Get
	End Property

	Public ReadOnly Property QuerySelectorAll(Query As String) As List(Of CEFElement)
		Get
			On Error Resume Next
			QuerySelectorAll = New List(Of CEFElement)
			Dim Code = "var a=document.all[{0}].querySelectorAll('{1}');var r='';for(i=0;i<a.length;i++){r+=a[i].srcIndex()+' '} return r"
			Dim Ret = CStr(Me.Document.Chrome.EvalScript(Code, SourceIndex, Query))
			If Ret.Length Then
				For Each Item In Ret.Split
					QuerySelectorAll.Add(Me.Document.All(Item))
				Next
			End If
		End Get
	End Property

	Public Property Style(Name As String) As String
		Get
			On Error Resume Next
			Dim Code = String.Format("return document.all[{0}].style.{1}", SourceIndex, Name)
			Return Me.Document.Chrome.EvalScript(Code)
		End Get
		Set(value As String)
			On Error Resume Next
			Dim Code = String.Format("document.all[{0}].style.{1}='{2}'", SourceIndex, Name, value)
			Me.Document.Chrome.ExecScript(Code)
		End Set
	End Property

	Public Shadows Function ToString() As String
		Return Me.TagName
	End Function

End Class

Public Class CEFMessage
	Event DOMEvent(Msg As String, Idx As Integer)

	Public Sub msg(Msg As String, Idx As Integer)
		RaiseEvent DOMEvent(Msg, Idx)
	End Sub

End Class

Public Class CEFKeyPreviewArgs
	Public Type As CefSharp.KeyType
	Public WindowsKeyCode As Integer
	Public NativeKeyCode As Integer
	Public Modifiers As CefEventFlags
	Public IsSystemKey As Boolean

End Class
