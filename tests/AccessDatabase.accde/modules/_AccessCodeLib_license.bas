Attribute VB_Name = "_AccessCodeLib_license"
'---------------------------------------------------------------------------------------
' access-codelib.net Lizenz
'---------------------------------------------------------------------------------------
'/**
' <summary>
' access-codelib.net Lizenz
' </summary>
' <remarks>
'---------------------------------------------------------------------------------------\n
' access-codelib.net Lizenz                                                             \n
'---------------------------------------------------------------------------------------\n
'
' Copyright (c) access-codelib.net
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without modification,
' are permitted provided that the following conditions are met:
'
' * Redistributions of source code must retain the above copyright notice,
'   this list of conditions and the following disclaimer.
' * Redistributions in binary form must reproduce the above copyright notice,
'   this list of conditions and the following disclaimer in the documentation
'   and/or other materials provided with the distribution.
' * Neither the name of access-codelib.net nor the names of its contributors may
'   be used to endorse or promote products derived from this software without specific
'   prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
' INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
' CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'---------------------------------------------------------------------------------------\n
' BSD-Lizenz im Originial: http://opensource.org/licenses/bsd-license.php               \n
'---------------------------------------------------------------------------------------\n
'
' Beachten Sie auch die Nutzungsbedingungen von access-codelib.net:
' http://access-codelib.net/nutzungsbedingungen.html
'
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/license.bas</file>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Public Function GetAccessCodeLibLicense() As String

On Error Resume Next

   GetAccessCodeLibLicense = _
      "Copyright (c) access-codelib.net" & vbNewLine & _
      "All rights reserved." & vbNewLine & vbNewLine & _
      "Redistribution and use in source and binary forms, with or without modification," & vbNewLine & _
      "are permitted provided that the following conditions are met:" & vbNewLine & _
      vbNewLine & _
      "* Redistributions of source code must retain the above copyright notice," & vbNewLine & _
      "  this list of conditions and the following disclaimer." & vbNewLine & _
      "* Redistributions in binary form must reproduce the above copyright notice," & vbNewLine & _
      "  this list of conditions and the following disclaimer in the documentation" & vbNewLine & _
      "  and/or other materials provided with the distribution." & vbNewLine & _
      "* Neither the name of access-codelib.net nor the names of its contributors may" & vbNewLine & _
      "  be used to endorse or promote products derived from this software without" & vbNewLine & _
      "  specific prior written permission." & vbNewLine & vbNewLine & _
      "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ""AS IS"" AND ANY EXPRESS OR IMPLIED WARRANTIES," & vbNewLine & _
      "INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE" & vbNewLine & _
      "DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL," & vbNewLine & _
      "SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;" & vbNewLine & _
      "LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN" & vbNewLine & _
      "CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS" & vbNewLine & _
      "SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."

End Function
