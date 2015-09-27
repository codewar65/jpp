Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports Microsoft.Ajax.Utilities

Module jpp

	'	code in #else and #elif
	'	write new eval to get rid of jscript sc.eval()

	'	#include "FILENAME"
	'	#define LVAL RVAL				
	'	#define MACRO([arg[,...]]) RVAL	
	'	#undef LVAL
	'	#ifdef LVAL
	'	#ifndef LVAL
	'	#if RVAL
	'	#endif 
	'	#else			(not implemented)
	'	#elif			(not implemented)
	'	#error STRING
	'	#pragma 
	'	#//
	'	__LINE__
	'	__FILE__
	'	__DATE__
	'	__TIME__

	Public Class progline
		' a line of source code
		Public lnum As Integer
		Public fname As String
		Public text As String
		Public cont As Boolean  ' this line continues on next.
		Sub New(_lnum As Integer, _fname As String, _text As String)
			lnum = _lnum
			fname = _fname
			text = _text
			cont = False
		End Sub
	End Class

	Public Class dictitem
		Public name As String
		Public val As String
		Public args() As String
		Sub New(_name As String, _val As String)
			name = _name
			val = _val
			args = Nothing
		End Sub
	End Class

	Public errmsgs() As String = {
		"Success",
		"Usage: jpp.exe [-o <output-file>] [<input-file>]",
		"User error",
		"Invalid directive",
		"Malformed filename",
		"Unable to find file",
		"Unable to open file",
		"Include file used already",
		"Endif without matching if",
		"Unclosed if block",
		"Expression error",
		"Can not modify built in value",
		"Macro syntax error",
		"Invalid number of arguments in macro",
		"Macro syntax error",
		"Macro too complex / excessive levels / recursion",
		"minifify error"
	}
	Const _ERR_SUCCESS As Integer = 0
	Const _ERR_HELP As Integer = 1
	Const _ERR_USER As Integer = 2
	Const _ERR_BADDIR As Integer = 3
	Const _ERR_BADFNAME As Integer = 4
	Const _ERR_FFOUND As Integer = 5
	Const _ERR_FOPEN As Integer = 6
	Const _ERR_REUSE As Integer = 7
	Const _ERR_ORPHANENDIF As Integer = 8
	Const _ERR_ORPHANIF As Integer = 9
	Const _ERR_EXPR As Integer = 10
	Const _ERR_BUILTIN As Integer = 11
	Const _ERR_BADMACRO As Integer = 12
	Const _ERR_NUMARGS As Integer = 13
	Const _ERR_BADMACRO2 As Integer = 14
	Const _ERR_COMPLEX As Integer = 15
	Const _ERR_MINI As Integer = 16

	Const _MAX_PARSELEVELS As Integer = 8

	Const _VARCHRS0 As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_$"
	Const _VARCHRSN As String = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_$"

	Public eol As String = vbNullChar
	Public skipstarts() As String = {"'", """", "/*", "//"}
	Public skipescs() As String = {"\'", "\""", Nothing, Nothing}
	Public skipterms() As String = {"'", """", "*/", eol}
	Public skipconts() As String = {"\" + eol, "\" & eol, Nothing, Nothing}
	Const _NOSKIP As Integer = -1
	Const _SQUOTE As Integer = 0
	Const _DQUOTE As Integer = 1
	Const _BLOCKCMT As Integer = 2
	Const _LINECMT As Integer = 3

	Public dict As New List(Of dictitem)
	Public sc As New MSScriptControl.ScriptControl()
	Public source As New List(Of progline)
	Public retcode As Integer = _ERR_SUCCESS
	Public tabstop As Integer = 8

	Public dirs() As String = {
		"#include", "#define", "#undef", "#ifdef", "#ifndef",
		"#if", "#endif", "#else", "#elif", "#error",
		"#pragma", "#//"
	}
	Const _INCLUDE = 0
	Const _DEFINE = 1
	Const _UNDEF = 2
	Const _IFDEF = 3
	Const _IFNDEF = 4
	Const _IF = 5
	Const _ENDIF = 6
	Const _ELSE = 7
	Const _ELIF = 8
	Const _ERROR = 9
	Const _PRAGMA = 10
	Const _COMMENT = 11

	Public parsecount As Integer = 0

	Function parse(str As String, Optional specvar As String = "", Optional specval As String = "") As String
		Dim n As String                 ' output string
		Dim v As String                 ' variable name
		Dim i As Integer                ' position idx
		Dim j As Integer
		Dim k As Integer
		Dim l As Integer                ' paren level
		Dim an As Integer               ' arg number
		Dim a As New List(Of String)    ' the args
		Dim d As Integer                ' dict idx
		Dim t As String                 ' macro building str
		Dim p As Integer
		Dim ic As String
		Dim jc As String
		Dim kc As String

		parsecount += 1
		If parsecount > _MAX_PARSELEVELS Then
			retcode = _ERR_COMPLEX
			Return Nothing
		End If
		n = ""
		If str.Substring(str.Length - 1, 1) <> vbNullChar Then
			str &= vbNullChar
		End If

		' temporary add spec to dictionary if available
		If specvar <> "" Then
			dict.Add(New dictitem(specvar, specval))
		End If
		For i = 0 To str.Length - 1
			v = ""
			ic = str.Substring(i, 1)
			If _VARCHRS0.IndexOf(ic) <> -1 Then
				' start of var name
				v += ic
				For j = i + 1 To str.Length - 1
					jc = str.Substring(j, 1)
					' get rest of var name
					If _VARCHRSN.IndexOf(jc) <> -1 Then
						v += jc
					Else
						If jc = "(" Then
							' macro - get args
							an = 0
							a.Clear()
							a.Add("")
							l = 1
							k = j + 1
							Do
								kc = str.Substring(k, 1)
								If k >= str.Length Then
									retcode = _ERR_BADMACRO2
									Return Nothing
								End If
								If kc = "(" Then
									l += 1
									a(an) &= kc
								ElseIf kc = ")" Then
									l -= 1
									If l = 0 Then
										Exit Do
									End If
									a(an) &= kc
								ElseIf kc = "," Then
									an += 1
									a.Add("")
								Else
									a(an) &= kc
								End If
								k += 1
							Loop
							' do str replacement here and reparse
							d = dictidx(v)
							If d = -1 Then
								' not in dictionary. verbatum and forward to first paren
								n &= v
								i = j - 1
								Exit For
							Else
								' replace this with dictionary item
								If dict(d).args.Count <> a.Count Then
									retcode = _ERR_BADMACRO2
									Return Nothing
								End If

								' parse each of the individual args
								For p = 0 To a.Count - 1
									a(p) = parse(a(p))
									If a(p) Is Nothing Then
										' back out on error
										Return Nothing
									End If
								Next

								' parse the whole macro replacing args
								t = dict(d).val
								For p = 0 To dict(d).args.Count - 1
									t = parse(t, dict(d).args(p), a(p))
									If t Is Nothing Then
										' back out on error
										Return Nothing
									End If
								Next
								n &= t
								i = k
								Exit For
							End If
						Else
							' simple replacement
							d = dictidx(v)
							If d = -1 Then
								' not in dictionary. verbatum
								n &= v
							Else
								n &= dict(d).val
							End If
							i = j - 1
							Exit For
						End If
					End If
				Next
			Else
				n &= ic
			End If
		Next

		' remove spec from dict
		If specvar <> "" Then
			dict.RemoveAt(dict.Count - 1)
		End If

		' remove eol marker
		parsecount -= 1
		If n.Substring(n.Length - 1, 1) = vbNullChar Then
			n = n.Substring(0, n.Length - 1)
		End If

		' try to reduce
		Try
			If n.IndexOf("""") = -1 And n.IndexOf("'") = -1 Then
				Dim rn As String
				rn = sc.Eval(n)
				If Not rn Is Nothing Then
					n = rn
				End If
			End If
		Catch ex As Exception
			' unsuccess. leave alone
		End Try
		Return n
	End Function

	Function readfilelines(ByVal fname As String) As List(Of progline)
		' read a file. return as array of lines or nothing if error
		' if fname is "", use stdin
		Dim sr As StreamReader
		Dim i As Integer = 0
		Dim litem As progline
		Dim source As New List(Of progline)

		If fname <> "" Then
			Try
				sr = New StreamReader(fname)
			Catch ex As Exception
				retcode = _ERR_FFOUND
				Return Nothing
			End Try
		Else
			sr = New StreamReader(Console.OpenStandardInput())
		End If
		Do While sr.Peek >= 0
			litem = New progline(i + 1, fname, sr.ReadLine)
			litem.text = litem.text.TrimEnd
			source.Add(litem)
			i += 1
		Loop
		sr.Close()
		Return source
	End Function

	Function dictidx(name As String) As Integer
		' return index number into dict of name or -1 if not found
		Dim i As Integer
		For i = 0 To dict.Count - 1
			If name = dict(i).name Then
				Return i
			End If
		Next
		Return -1
	End Function

	Function dictval(name As String, slnum As Integer) As String
		' return value in dictionary or nothing if not found
		Dim i As Integer
		For i = 0 To dict.Count - 1
			If name = dict(i).name Then
				Select Case name
					Case "__LINE__"
						Return source(slnum).lnum.ToString

					Case "__FILE__"
						Return """" & source(slnum).fname & """"

					Case "__TIME__"
						Return """" & Now.TimeOfDay.ToString() & """"

					Case "__DATE__"
						Return """" & Now.Date.ToString() & """"

					Case Else
						Return dict(i).val
				End Select
			End If
		Next
		Return Nothing
	End Function

	Function Main(ByVal args() As String) As Integer
		Dim infname As String
		Dim outfname As String
		Dim inblockcmt As Boolean = False   ' inside a comment block
		Dim templ As String
		Dim lineargs() As String
		Dim slnum As Integer = 0
		Dim i As Integer = 0
		Dim sw As StreamWriter
		Dim dir As Integer
		Dim found As Boolean
		Dim ifname As String                ' include filename
		Dim lval As String
		Dim rval As String
		Dim val As String
		Dim inif As Boolean = False
		Dim ifstate As Boolean = True       ' true=include,false=skip
		Dim iflevel As Integer = 0
		Dim ifskip As Boolean = False
		Dim genstack As New Stack
		Dim stripemptylines As Boolean = False  ' pragma setting
		Dim stripcomments As Boolean = False    ' pragma setting
		Dim addfrom As Boolean = False          ' pragma setting
		Dim minify As Boolean = False           ' pragma setting (won't work with addfrom or strip pragmas)
		Dim pack As Boolean = False
		Dim skipstate As Integer = _NOSKIP
		Dim k As Integer
		Dim l As Integer
		Dim kc As String
		Dim ifiles As New List(Of String)
		Dim maxlinelen As Integer
		Dim isource As List(Of progline)
		Dim p1 As Integer
		Dim p2 As Integer
		Dim macroargstr As String
		Dim macroargs() As String
		Dim di As dictitem
		Dim nc As String
		Dim pc As String
		Dim newstr As String
		Dim parseval As String
		Dim newlen As Integer
		Dim minifier = New Minifier
		Dim skipstart As Integer
		Dim packstr As String
		Dim c As String

		infname = ""
		outfname = ""
		For i = 0 To args.Count - 1
			Select Case args(i)
				Case "-?", "/?"
					Console.Error.WriteLine("jpp - JavaScript Precompiler" & vbCrLf & "(c)2015  Dan Mecklenburg Jr  v0.1" & vbCrLf)
					retcode = _ERR_HELP
					GoTo done

				Case "-o", "-O", "/o", "/O"
					i += 1
					outfname = args(i)

				Case Else
					infname = args(i)
			End Select
		Next
		source = readfilelines(infname)
		If source Is Nothing Then
			retcode = _ERR_FOPEN
			GoTo done
		End If

		' keep track of files we've included so we can't include them again.

		dict.Add(New dictitem("__LINE__", Nothing))
		dict.Add(New dictitem("__FILE__", Nothing))
		dict.Add(New dictitem("__TIME__", Nothing))
		dict.Add(New dictitem("__DATE__", Nothing))

		' walk through lines of file.
		sc.Language = "JScript"
		maxlinelen = 0
		While slnum < source.Count
			skipstart = skipstate
			templ = source(slnum).text.Trim
			templ = Regex.Replace(templ, "\s+", " ")
			found = False

			For dir = 0 To dirs.Length - 1
				If templ.Length >= dirs(dir).Length Then
					If templ.Substring(0, dirs(dir).Length) = dirs(dir) Then
						found = True
						Exit For
					End If
				End If
			Next

			If found Then
				lineargs = templ.Split(" ")
				Select Case dir
					Case _INCLUDE
						If Not inif Or ifstate Then
							' fetch lines for include file
							' remove "'s from filename
							ifname = lineargs(1).Trim
							If ifname.Length < 2 Then
								retcode = _ERR_BADFNAME
								GoTo done
							End If
							If ifname.Substring(0, 1) <> """" Or ifname.Substring(ifname.Length - 1, 1) <> """" Then
								retcode = _ERR_BADFNAME
								GoTo done
							End If
							ifname = ifname.Substring(1, ifname.Length - 2)

							' check if it's been used
							For i = 0 To ifiles.Count - 1
								If ifiles(i) = ifname Then
									retcode = _ERR_REUSE
									GoTo done
								End If
							Next

							ifiles.Add(ifname)
							isource = New List(Of progline)
							isource = readfilelines(ifname)
							If isource Is Nothing Then
								retcode = _ERR_FOPEN
								GoTo done
							End If

							' insert into array
							source.RemoveAt(slnum)
							For i = isource.Count - 1 To 0 Step -1
								source.Insert(slnum, isource(i))
							Next
						Else
							source.RemoveAt(slnum)
						End If

					Case _DEFINE
						If Not inif Or ifstate Then
							lval = lineargs(1)

							' if "(args [, ...])" after lval, then macro
							p1 = lval.IndexOf("("c)
							macroargstr = ""
							macroargs = Nothing
							If p1 <> -1 Then
								' macro
								p2 = lval.IndexOf(")"c)
								If p2 = -1 Then
									' no closing paren
									retcode = _ERR_BADMACRO
									GoTo done
								End If
								' extract args
								macroargstr = lval.Substring(p1 + 1, p2 - p1 - 1)
								macroargs = macroargstr.Split(","c)
								lval = lval.Substring(0, p1)
							End If

							rval = ""
							For i = 2 To lineargs.Length - 1
								rval &= lineargs(i) & " "
							Next
							rval = rval.Trim
							If rval = "" Then
								rval = "1"
							End If

							' add / update dictionary lval = val
							i = dictidx(lval)
							If i <> -1 Then
								' redefine
								If dict(i).val Is Nothing Then
									' null val means build in var
									retcode = _ERR_BUILTIN
									GoTo done
								Else
									rval = parse(rval)
									dict(i).val = rval
									dict(i).args = macroargs
								End If
							Else
								' new dict item
								rval = parse(rval)
								di = New dictitem(lval, rval)
								di.args = macroargs
								dict.Add(di)
							End If
						End If
						source.RemoveAt(slnum)

					Case _UNDEF
						If Not inif Or ifstate Then
							' delete dict item
							lval = lineargs(1)
							i = dictidx(lval)
							If i <> -1 Then
								If dict(i).val Is Nothing Then
									retcode = _ERR_BUILTIN
									GoTo done
								Else
									dict.RemoveAt(i)
								End If
							End If
						End If
						source.RemoveAt(slnum)

					Case _IFDEF
						' evaluate
						'If inif And Not ifstate Then Exit Select
						genstack.Push(ifstate)
						If Not inif Or ifstate Then
							inif = True
							If dictidx(lineargs(1)) <> -1 Then
								ifstate = True
							Else
								ifstate = False
							End If
						End If
						source.RemoveAt(slnum)

					Case _IFNDEF
						genstack.Push(ifstate)
						If Not inif Or ifstate Then
							inif = True
							If dictidx(lineargs(1)) <> -1 Then
								ifstate = False
							Else
								ifstate = True
							End If
						End If
						source.RemoveAt(slnum)

					Case _IF
						genstack.Push(ifstate)
						If Not inif Or ifstate Then
							inif = True
							' eval here
							rval = ""
							For i = 1 To lineargs.Length - 1
								rval &= lineargs(i)
							Next
							rval = parse(rval)
							'rval = normrval(rval, slnum)
							If rval Is Nothing Then
								retcode = _ERR_EXPR
								GoTo done
							End If
							val = sc.Eval(rval)
							If val <> "0" And val <> "False" Then
								ifstate = True
							Else
								ifstate = False
							End If
						End If
						source.RemoveAt(slnum)

					Case _ENDIF
						If Not inif Then
							retcode = _ERR_ORPHANENDIF
							GoTo done
						End If
						ifstate = genstack.Pop
						If genstack.Count = 0 Then
							inif = False
						End If
						source.RemoveAt(slnum)

					Case _ERROR
						If Not inif Or ifstate Then
							rval = ""
							For i = 1 To lineargs.Length - 1
								rval &= lineargs(i) & " "
							Next
							retcode = _ERR_USER
							errmsgs(_ERR_USER) = rval
							GoTo done
						End If
						source.RemoveAt(slnum)

					Case _PRAGMA
						If Not inif Or ifstate Then
							For i = 1 To lineargs.Length - 1
								Select Case lineargs(i).ToLower
									Case "stripemptylines"
										' remove empty lines from output
										stripemptylines = True

									Case "stripcomments"
										' remove comments from output
										stripcomments = True

									Case "addfrom"
										' add comment to end of lines of source reference
										addfrom = True

									Case "tab2"
										' set tabstops for formating addfrom
										tabstop = 2
									Case "tab3"
										tabstop = 3
									Case "tab4"
										tabstop = 4
									Case "tab5"
										tabstop = 5
									Case "tab6"
										tabstop = 6
									Case "tab7"
										tabstop = 7
									Case "tab8"
										tabstop = 8

									Case "minify"
										' run AjaxMini on output
										minify = True

									Case "pack"
										pack = True

								End Select
							Next
						End If
						source.RemoveAt(slnum)

					Case _COMMENT
						source.RemoveAt(slnum)

				End Select
			Else
				newstr = ""
				templ = source(slnum).text & vbNullChar
				i = 0
				Do
					If i >= templ.Length Then
						Exit Do
					End If

					If skipstate <> _NOSKIP Then

						' look for escape to skip
						If Not skipescs(skipstate) Is Nothing Then
							If i + skipescs(skipstate).Length <= templ.Length Then
								If templ.Substring(i, skipescs(skipstate).Length) = skipescs(skipstate) Then
									' escaped - pass over
									If Not stripcomments Or (skipstarts(skipstate) <> "//" And skipstarts(skipstate) <> "/*") Then
										newstr &= skipescs(skipstate)
									End If
									i = i + skipescs(skipstate).Length
									GoTo doneskip
								End If
							End If
						End If

						' look for terminator
						If i + skipterms(skipstate).Length <= templ.Length Then
							If templ.Substring(i, skipterms(skipstate).Length) = skipterms(skipstate) Then
								' terminator
								If Not stripcomments Or (skipstarts(skipstate) <> "//" And skipstarts(skipstate) <> "/*") Then
									newstr &= skipterms(skipstate)
								Else
									If skipterms(skipstate).Substring(skipterms(skipstate).Length - 1, 1) = vbNullChar Then
										newstr &= vbNullChar
									End If
								End If
								i = i + skipterms(skipstate).Length
								skipstate = _NOSKIP
								GoTo doneskip
							End If
						End If

						' look for continuation
						If Not skipconts(skipstate) Is Nothing Then
							If i + skipconts(skipstate).Length <= templ.Length Then
								If templ.Substring(i, skipconts(skipstate).Length) = skipconts(skipstate) Then
									source(slnum).cont = True
									If Not stripcomments Or (skipstarts(skipstate) <> "//" And skipstarts(skipstate) <> "/*") Then
										newstr &= skipconts(skipstate)
									End If
									i = i + skipconts(skipstate).Length
									GoTo doneskip
								End If
							End If
						End If

						' copy verbatum (unless we are in comment and stripcomments on)
						If Not stripcomments Or (skipstarts(skipstate) <> "//" And skipstarts(skipstate) <> "/*") Then
							newstr &= templ.Substring(i, 1)
						End If
						i += 1
doneskip:
					Else
						' look for start of skip
						For j = 0 To skipstarts.Length - 1
							If i + skipstarts(j).Length <= templ.Length Then
								If templ.Substring(i, skipstarts(j).Length) = skipstarts(j) Then
									If Not stripcomments Or (skipstarts(j) <> "//" And skipstarts(j) <> "/*") Then
										newstr &= skipstarts(j)
									End If
									i = i + skipstarts(j).Length
									skipstate = j
									GoTo doneskip2
								End If
							End If
						Next

						' no skip found. look for dict item to replace
						For j = 0 To dict.Count - 1
							If i + dict(j).name.Length <= templ.Length Then
								If templ.Substring(i, dict(j).name.Length) = dict(j).name Then
									' next character after this must be a non-idententifer
									p1 = i + dict(j).name.Length
									nc = templ.Substring(p1, 1)
									If _VARCHRSN.IndexOf(nc) = -1 Then
										' check for macro (arg[,...])
										If nc <> "(" Then
											' straight replace
											parsecount = 0
											parseval = parse(dict(j).val)
											If parsecount > _MAX_PARSELEVELS Then
												retcode = _ERR_COMPLEX
												GoTo done
											End If
											newstr &= parseval
											i = i + dict(j).name.Length
											GoTo doneskip2
										Else
											' get string of name and args for parse()
											'find matching )
											l = 1
											k = p1 + 1
											Do
												kc = templ.Substring(k, 1)
												If k >= templ.Length Then
													retcode = _ERR_BADMACRO2
													GoTo done
												End If
												If kc = "(" Then
													l += 1
												ElseIf kc = ")" Then
													l -= 1
													If l = 0 Then
														Exit Do
													End If
												End If
												k += 1
											Loop
											' k points to matching )
											parseval = parse(templ.Substring(i, k - i + 1))
											If parseval Is Nothing Then
												retcode = _ERR_COMPLEX
												GoTo done
											End If
											newstr &= parseval
											i = k + 1
											GoTo doneskip2
										End If
									End If
								End If
							End If
						Next

						' output verbatum
						Dim c1 As String = templ.Substring(i, 1)
						newstr &= c1
						i += 1
doneskip2:
					End If
				Loop

				If templ.Length > 0 Then
					If templ.Substring(0, 1) = "#" Then
						retcode = _ERR_BADDIR
						GoTo done
					End If
				End If

				If Not ifstate Then
					source.RemoveAt(slnum)
				Else
					' remove nul on end
					If newstr.Length > 0 Then
						If (newstr.Substring(newstr.Length - 1, 1)) = vbNullChar Then
							newstr = newstr.Substring(0, newstr.Length - 1)
						End If
					End If

					' pack here
					If pack Then
						Dim skipstatex As Integer
						i = 0
						packstr = ""
						skipstatex = skipstart
						Do
							If i = newstr.Length Then
								Exit Do
							End If
							c = newstr.Substring(i, 1)
							If skipstatex <> _NOSKIP Then
								' in a state of skip
								If Not skipescs(skipstatex) Is Nothing Then
									If i + skipescs(skipstatex).Length < newstr.Length - 1 Then
										If newstr.Substring(i, skipescs(skipstatex).Length) = skipescs(skipstatex) Then
											' escape. ignore
											packstr &= skipescs(skipstatex)
											i += skipescs(skipstatex).Length
											GoTo packskip
										End If
									End If
								End If

								If Not skipterms(skipstatex) Is Nothing Then
									If i + skipterms(skipstatex).Length < newstr.Length - 1 Then
										If newstr.Substring(i, skipterms(skipstatex).Length) = skipterms(skipstatex) Then
											' terminator
											packstr &= skipterms(skipstatex)
											i += skipterms(skipstatex).Length
											skipstatex = _NOSKIP
											GoTo packskip
										End If
									End If
								End If

								packstr &= c
								i += 1
							Else
								' not skipping
								' search for skipstart
								If Asc(c) > 32 Then
									For j = 0 To skipstarts.Count - 1
										If i + skipstarts(j).Length < newstr.Length Then
											If newstr.Substring(i, skipstarts(j).Length) = skipstarts(j) Then
												' start a skip
												packstr &= skipstarts(j)
												i += skipstarts(j).Length
												skipstatex = j
												GoTo packskip
											End If
										End If
									Next
								Else
									c = " "
									pc = " "
									If i > 0 Then
										pc = newstr.Substring(i - 1, 1)
									End If
									nc = " "
									If i + 1 < newstr.Length Then
										nc = newstr.Substring(i + 1, 1)
									End If
									If pc = " " Then
										i += 1
										GoTo packskip
									End If
									If (pc = "+" Or pc = "-") And (nc = "+" Or nc = "-") Then
										GoTo packnoskip
									End If
									If _VARCHRSN.IndexOf(pc) = -1 Or _VARCHRSN.IndexOf(nc) = -1 Then
										i += 1
										GoTo packskip
									End If
								End If
packnoskip:                     packstr &= c
								i += 1
							End If
packskip:
						Loop
						newstr = packstr
					End If

					' figure maxlinelen
					If getlinelen(newstr) > maxlinelen Then
						maxlinelen = getlinelen(newstr)
					End If
					source(slnum).text = newstr
					slnum += 1
				End If
			End If
		End While

		If genstack.Count > 0 Then
			retcode = _ERR_ORPHANIF
			GoTo done
		End If

		' write result
		If minify Then
			Dim mini As String = ""
			For i = 0 To source.Count - 1
				mini &= source(i).text & vbCrLf
			Next
			If outfname <> "" Then
				sw = New StreamWriter(outfname)
			Else
				sw = New StreamWriter(Console.OpenStandardOutput())
			End If
			sw.Write(Minifier.MinifyJavaScript(mini))
		Else
			newlen = (Int(maxlinelen / tabstop) + 1) * tabstop
			If outfname <> "" Then
				sw = New StreamWriter(outfname)
			Else
				sw = New StreamWriter(Console.OpenStandardOutput())
			End If
			For i = 0 To source.Count - 1
				If Not stripemptylines Or source(i).text.Trim.Length > 0 Then
					If addfrom And Not source(i).cont Then
						sw.WriteLine(source(i).text & tabtopos(getlinelen(source(i).text), newlen) & "// " & source(i).fname & ":" & source(i).lnum.ToString & " ")
					Else
						sw.WriteLine(source(i).text)
					End If
				End If
			Next
		End If
		sw.Close()
done:
		If retcode > 0 Then
			Console.Error.WriteLine("Error " & retcode & ": " & errmsgs(retcode))
			If Not source Is Nothing Then
				If (slnum <= source.Count - 1) Then
					Console.Error.WriteLine("File: " & source(slnum).fname & " Line: " & source(slnum).lnum)
				End If
			End If
		Else
			Console.Error.WriteLine(errmsgs(retcode))
		End If
		Return retcode
	End Function

	Function tabtopos(fromcol As Integer, tocol As Integer) As String
		Dim tstr As String = ""

		While fromcol < tocol
			tstr &= vbTab
			fromcol = (Int(fromcol / tabstop) + 1) * tabstop
		End While
		Return tstr
	End Function

	Function getlinelen(str As String) As Integer
		' return position at end of string based on tabstop
		Dim col As Integer

		col = 0
		For i = 0 To str.Length - 1
			If str.Substring(i, 1) = vbTab Then
				col = (Int(col / tabstop) + 1) * tabstop
			Else
				col = col + 1
			End If
		Next
		Return col
	End Function

End Module
