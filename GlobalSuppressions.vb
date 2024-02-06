' This file is used by Code Analysis to maintain SuppressMessage
' attributes that are applied to this project.
' Project-level suppressions either have no target or are given
' a specific target and scoped to a namespace, type, member, etc.

Imports System.Diagnostics.CodeAnalysis

<Assembly: SuppressMessage("Minor Code Smell", "S1643:Strings should not be concatenated using ""+"" or ""&"" in a loop", Justification:="<Pending>", Scope:="member", Target:="~M:libBAUtil.LoggingHelper.GetMethodParameters(System.Reflection.MethodBase,System.Object[])~System.String")>
<Assembly: SuppressMessage("Major Code Smell", "S6562:Always set the ""DateTimeKind"" when creating new ""DateTime"" instances", Justification:="<Pending>", Scope:="member", Target:="~M:libBAUtil.DateTimeHelper.#ctor(System.Int32,System.Int32,System.Int32,System.Int32,System.Int32,System.Int32,System.Globalization.Calendar)")>
