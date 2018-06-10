#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=StringElaborate.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=String functions
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 2018-06-08
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4 -w 5
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 2 /reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /rm
#Au3Stripper_Ignore_Funcs=StringTitleCase, StringFloat
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; Tested with: AutoIt version 3.3.14.5 and win10 / win7

#include-once
#include <EditConstants.au3> ; Delivered with AutoIT
#include <GuiEdit.au3> ; Delivered with AutoIT
#include <GUIConstantsEx.au3> ; Delivered with AutoIT
#include <WindowsConstants.au3> ; Delivered with AutoIT
#include <aDrosteArray.au3> ; Published at GitHub
#include <FileJuggler.au3> ; Published at GitHub

#Region ; Regulated expression interface with procedural syntax ===============

; Advanced search
; As more and more digital texts pile up, the need to find a needle in a
; haystack increases. With regulated expressions (RegEx) one can effectively
; do an advanced search in large amounts of text files and find that needle.
;
; Structure informal data
; Lots of more or less structured data appear. Some people see a pattern and
; may want to get things in order. RegEx can transform stuff, bring things
; into a format suitable for further computer processing.
;
; Reformat between applications
; Many times data needs a transformation into a format suitable for another
; applicaton. RexEx can do just that.
;
; Speed
; The RegEx technique works with interpreter programming packages as well as
; with compilers. Because a RegEx module is a kind of compiler in it self,
; it will do tiny character byte operations in microseconds where procedural
; syntax of interpreter progamming is slower. That is why RegEx is not to be
; missed by programmers concerned with textfiles.
;
; Career
; Computer programming languages differ in many respects. Programmers that
; learn to use RegEx will see that the syntax and the working is much alike on
; many platforms.

; Standardization
; Many instant data formats arised out of necesity, without much inter-toko
; discussion. For example movie subtitle formats such as .srt .sub and
; .ass did see the light, for players it is a maze what exactly to expect.
; With RegEx pattern script, formats can be described precisely. A formal
; description will make data more interchangable between applications. Please
; communicate your experiences with patterns.
;
; StringsElaborate.au3
; This source offers several functions to work with RegEx in AutoIT. They all
; have 'RegEx in their name. One can derive the type the value they return from
; the first letter:
;
;     i  denotes a number pointing to the input string
;     s  stands for string
;     a  these functions return an array
;     m  RegEx properties are returned: an array with named parameters
;     g  returns an array with mRegEx maps
;
; The first five functions in the list below search for the first match. The
; next ones can find multiple matches. For ease of use many functions have the
; same input parameters:
;
; Function        Input-parameters                        Return value
; -----------------------------------------------------------------------------
; iRegExCore()    $sInput $sPattern $iOffset $aCapture    $iFound
; mRegEx()        $sInput $sPattern $iStart               Array of $mProperties
; iRegEx()        $sInput $sPattern $iStart               $iFound
; sRegEx()        $sInput $sPattern $iStart  $sSeparator  $sFound
; aRegEx()        $sInput $sPattern $iStart               $aCaptures
; gRegEx()        $sInput $sPattern $iStart               Array of $mRegex maps
; aRegExLocate()  $sInput $sPattern $iStart  $Last        $aLocations
; aRegexFindAll() $sInput $sPattern $iStart  $sSeparator  $aMatches
; aRegExSplit()   $sInput $sPattern $iStart               $sParts
; sRegExReplace() $sInput $sPattern $sScript $iStart      $sOut

; The first two input parameters $sInput and $sPattern are the same for all
; functions. From the third parameter on, the parameters are optional. There is
; one exception: iRegExCore() always needs four parameters. This function
; passes $aCapture and $iFound by reference. They receive output from the
; iRegExCore function. So one needs to declare these variables beforehand.
;
; Some of the functions return a value via macro @extended. However there are
; notable differences in what the value means:
;
;    + With iRegExCore() sRegEx() and $aRegex(), @extended return the start
;      location of the match in the input string.
;    + After execution of aRegExFindAll() and aRegExLocate(), @extended points
;      to the location directly after the match.
;    + Finally sRegExReplace() delivers in @extended the number of replacements
;      made.
;
; To help learn RegEx patterns, these Regex functions issue an informative halt
; message to the console if a pattern has a syntax error. @error code 2 is
; immediately intercepted within the function and brought to the attention of
; the programmer. So the programmer never has to bother about further
; processing that would obscure the cause. This is a time saver.
;
; The functions never raise the @error flag also if there is no match. Some
; functions don't return matches at all. With other functions the iFound
; property signals wether there is a match. For functions that process multiple
; matches, see the count property of the returned array. Find the iFound
; property of sRegEx() and aRegEx() in the macro @extended.
;
; Patterns
; To practice with the formation of patterns, AutoIT offers a helper GUI. Find
; it by searching on: "StringRegEpxGUI in the help files, then browse the page
; to "Resources" under "Example 6".
;
; Excellent documentation is at: "http://regular-expressions.info"
; and in: "Regular Expressions Cookbook" written by Jan Goyvaerts and Steven
; Levithan. The hard-copy is published by O'Reilly and available via Amazon.
;
; Unicode
; All the functions in this interface operate on utf input strings by default.
; To comply, the first thing to do is: save your code with your source code
; editor in utf8 format. For old-style ascii input, precede the pattern with
; syntax: (*ASC)
;
; I hope the functions presented here will facilitate the use of RegEx in
; AutoIT and make people curious about regulated expressions and what they can
; do for you.
;
; Author ....: Bert Kerkhof ( kerkhof.bert@gmail.com )

Func iRegExCore(Const $sInput, Const $sPattern, ByRef $iOffset, ByRef $aCapture)
  ; Ten Features That Makes This RegEx Interface Excellent:
  ; [1]  This core function search for a match and returns four major RegEx
  ;      data elements back to the user.
  ; [2]  Since a function can return only one value, two elements in the
  ;      parameter-list are passed by reference. The user has to declare the
  ;      variables $aCapture and $iOffset beforehand. Novice users may be
  ;      unacquainted with the pass by reference technique. That is why there
  ;      are other functions that directly return the number of matches, or the
  ;      full match string, or captured fields within the match, thereby
  ;      surpassing other information elements.
  ; [3]  iRegExCore supports an option to capture fields, these are fragments
  ;      within the match. The function returns the captures in an array
  ;      $aCapture.
  ; [4]  The element numbers in the $aCapture array correspond with the RegEx
  ;      numbering of captures.
  ; [5]  If there are no capture fields, $aCapture element number one captures
  ;      the full match.
  ; [6]  Find the total number of captures in the zero element of $aCapture.
  ; [7]  If there is no match, $aCapture is still an array with zero element.
  ; [8]  All returned elements are always of the expected data-type. This
  ;      feature prevents the use of error checking and is missed in many other
  ;      RegEx implementations.
  ; [9]  Finding the right RegEx pattern may be difficult for beginners, but
  ;      the absence of any exceptional behavior of returned elements will
  ;      facilitate the journey. It will lead to [A] a less steep learning
  ;      curve with RegEx patterns, [B] easy going programming and [C]
  ;      sustainable software development.
  ; [10] With the $iOffset variable, the RegEx search for a match can start on
  ;      a position in the string other than 1. On return, $iOffset contains
  ;      the first location after the match. This feature is designed for use
  ;      of the iRegExCore function in iterations where data are gathered from
  ;      multiple matches.
  ;
  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The search criterion.
  ;               + $iOffset.... Start-location in the string where the search
  ;                              will start.
  ;               + $aCapture... An array parameter, usually declared as:
  ;                                  $aCapture=aNewArray(), same as:
  ;                                  Local $aCapture[1]=[0]
  ; Returns:..... + The location found of the match. This is the first position
  ;                 of the match in the input string. If there is no match,
  ;                 zero is returned.
  ;               + Passed by reference parameter $aCapture contains the
  ;                 captured fields as an aDroste array of strings which
  ;                 means: the zero element of the array contains the number
  ;                 of captured fields.
  ;               + $iOffset points to the first location after the match.
  ; AutoIT bonus: + Macro @extended points to the start location of the match.
  ;                 If there is no match, @extended is zero.
  ;
  ; Warnings:
  ; [1] Any content in $aCapture is overwritten and $iOffset has no build-in
  ;     default value.
  ; [2] These parameters are not optional.

  Local $sPat = StringLeft($sInput, 6) = "(*ASC)" ? "" : "(*UTF)" & $sPattern
  $aCapture = StringRegExp($sInput, $sPat, $STR_REGEXPARRAYFULLMATCH, $iOffset)
  Local $error = @error, $extended = @extended
  If $error = 2 Then ; error in pattern
    cc('RegEx pattern error: ' & $sPattern & @CRLF & Spaces($extended + 15) & "^")
    Exit ; fatal
  EndIf
  If $error Or Not IsArray($aCapture) Then
    $aCapture = aNewArray()
    Return SetError(0, $extended, 0)
  EndIf
  $iOffset = $extended
  $extended -= StringLen($aCapture[0]) ; Calculate start-location of the full match
  If UBound($aCapture) > 1 Then
    $aCapture[0] = UBound($aCapture) - 1
  Else
    $aCapture = aConcat($aCapture[0])
  EndIf
  Return SetError(0, 0, $extended)
EndFunc   ;==>iRegExCore

#EndRegion ; Regulated expression interface with procedural syntax ===============

#Region ; Regulated expressions interface with mapped properties ==============

; Mapped properties in array of type mRegEx:
Global Enum $rSTART = 1, $rFOUND, $rOFFSET, $rCAPTURE

Func mRegExdotInit($iStart)
  ; This is a method that belongs to mRegex()
  ; Init provides for default property values in case a RegEx has no match.
  ; Thanks w0uter
  Return aConcat($iStart, 0, $iStart, aNewArray(0))
EndFunc   ;==>mRegExdotInit

Func mRegExdotHeader(Const $sInput, Const $mRegEx)
  ; This is a method that belongs to mRegex()
  ; Header is the text starting at .iStart upto .iFound
  Return StringMid($sInput, $mRegEx[$rSTART], $mRegEx[$rFOUND] - $mRegEx[$rSTART])
EndFunc   ;==>mRegExdotHeader

Func mRegExdotFooter(Const $sInput, $mRegEx)
  ; This is a method that belongs to mRegex()
  ; With N matches, there are N+1 parts in between. Thats where Footer comes in.
  ; Footer is the text starting at .iOffset up to the end of sInput.
  Return StringMid($sInput, $mRegEx[$rOFFSET])
EndFunc   ;==>mRegExdotFooter

Func mRegEx(Const $sInput, Const $sPattern, $iStart = 1)
  ; Features:
  ; [1] Finds a match and returns a map of RegEx properties.
  ; [2] The properties are named: iStart, iFound, aCapture, iOffset

  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The regex search criteria.
  ;               + $iStart..... [Optional] location in $sInput where the
  ;                              search starts. Default is 1.
  ; Returns:..... + An array with four named elements:
  ;                   .iStart    This number points to the first position
  ;                              in sInput where the search did begin.
  ;                   .iFound    Points to the fist position of the match
  ;                              in the sInput string. If there is no
  ;                              match, its value is zero.
  ;                   .aCapture  The elements in this array, starting at
  ;                              index 1, contain the captured fields. The
  ;                              index number is equal to the RegEx
  ;                              capture number. If the pattern does not
  ;                              provide for a capture, the first element
  ;                              presents the full match. Element 0 is the
  ;                              count property: the number of elements.
  ;                   .iOffset   After a match, this points to the first
  ;                              position in sInput directly after the match.
  ;               + If there is no match, still this array is returned with
  ;                 .iFound set to 0 and aCapture with count zero.
  ; Note:
  ; The property features used in this function have the form of named constants
  ; that refer to the index of the mRegEx array. To practice with another
  ; possible implementation, try AutoIT beta v3.3.13.21

  Local $mRegEx = mRegExdotInit($iStart)
  $mRegEx[$rFOUND] = iRegExCore($sInput, $sPattern, $mRegEx[$rOFFSET], $mRegEx[$rCAPTURE])
  Return $mRegEx
EndFunc   ;==>mRegEx

Func iRegEx(Const $sInput, Const $sPattern, $iStart = 1)
  ; Features:
  ; [1] Finds a match and returns the location in the string where the match
  ;     is found.
  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The regex search criteria.
  ;               + $iStart..... [Optional] location in $sInput where the
  ;                              search starts. Default is 1.
  ; Returns:..... + If there is a match, the function returns the location of
  ;                 the match found in the input string. This number can also
  ;                 be interpreted as logic (True). Otherwise returns 0 (False).

  Local $mRegEx = mRegEx($sInput, $sPattern, $iStart)
  Return SetError(0, 0, $mRegEx[$rFOUND])
EndFunc   ;==>iRegEx

Func sRegEx(Const $sInput, Const $sPattern, $iStart = 1, Const $sSep = "|")
  ; Features:
  ; [1] Finds a match and returns a string of captured elements.
  ; [2] If there is more than one capture, the elements are separated by a
  ;     vertical bar "|".
  ; [3] Optionally another character or string may be chosen as separator.

  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The regex search criteria.
  ;               + $iStart..... [Optional] location in $sInput where the
  ;                              search starts. Default is 1.
  ;               + $sSep....... [Optional] separates the values in the
  ;                              returned array, see features. Default is
  ;                              the vertical bar character '|'.
  ; Returns:..... + A string of captured elements.
  ; AutoIT bonus: + Macro @extended points to the start location of the match.
  ;                 If there is no match, @extended is zero.

  Local $mRegEx = mRegEx($sInput, $sPattern, $iStart)
  Return SetError(0, $mRegEx[$rFOUND], sRecite($mRegEx[$rCAPTURE], $sSep))
EndFunc   ;==>sRegEx

Func aRegEx(Const $sInput, Const $sPattern, $iStart = 1)
  ; Features:
  ; [1] Finds a match and returns an array of captured elements, even if the
  ;     pattern did not match.
  ; [2] Find the total number of captured fields in the zero element of the
  ;     array.

  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The regex search criteria.
  ;               + $iStart..... [Optional] location in $sInput where the
  ;                              search starts. Default is 1.
  ; Returns:..... + An aDroste array of strings, the captured elements.
  ; AutoIT bonus: + Macro @extended points to the start location of the match.
  ;                 If there is no match, @extended is zero.

  Local $mRegEx = mRegEx($sInput, $sPattern, $iStart)
  Return SetError(0, $mRegEx[$rFOUND], $mRegEx[$rCAPTURE])
EndFunc   ;==>aRegEx

Func gRegEx(Const $sInput, Const $sPattern, $iStart = 1)
  ; Features:
  ; [1] Finds multiple matches.
  ; [2] The gRegEx array together with the msBrace() adaptation is able to
  ;     function as a collection in am iterative statement like:
  ;     For $mRegex In msBrace(gRegEx())
  ;
  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The RegEx search criteria.
  ;               + $iStart..... [Optional] location in $sInput where
  ;                              the search starts. Default is 1.
  ; Returns:..... A series of mRegEx elements in an array that starts
  ;               at index position zero with a count of elements.
  ;               Each mRegex element offers the same properties
  ;               as function mRegEx(). They are named: [1] iStart,
  ;               [2] iFound, [3] aCapture and [4] iOffset

  Local $gRegEx = aNewArray(0), $mRegEx = mRegExdotInit($iStart)
  While True
    $mRegEx[$rFOUND] = iRegExCore($sInput, $sPattern, $mRegEx[$rOFFSET], $mRegEx[$rCAPTURE])
    If $mRegEx[$rFOUND] = 0 Then ExitLoop
    aAdd($gRegEx, $mRegEx)
    $mRegEx[$rSTART] = $mRegEx[$rOFFSET]
  WEnd
  Return $gRegEx
EndFunc   ;==>gRegEx

Func aRegExLocate($sInput, $sPattern, $iStart = 1, $Last = False)
  ; Features:
  ; [1] Returns a one-dimensional array with multiple findings.
  ; [2] The total number of matches is in the zero element.
  ; [3] Other elements in the array contain the start locations of matches.

  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The RegEx search condition.
  ;               + $iStart..... (Optional) location in $sInput where search
  ;                              starts.
  ;               + $Last....... (Optional) If False (which is the default),
  ;                              the numbers in the returned array are pointers
  ;                              to the location of the first char of the match.
  ;                              If True, the numbers point to the last char
  ;                              of the match.
  ; Returns:..... + A aDroste array of numbers which point to the locations of
  ;                 the full matches.
  ; AutoIT bonus: + Macro @extended points to the location after the last
  ;                 match, so can be used to detect a footer in the input
  ;                 source.

  Local $aR = aNewArray(0), $mRegEx = mRegExdotInit($iStart)
  Local $I, $gRegEx = gRegEx($sInput, $sPattern, $iStart)
  For $I = 1 To $gRegEx[0]
    $mRegEx = $gRegEx[$I]
    Local $iPos = $Last ? $mRegEx[$rOFFSET] - 1 : $mRegEx[$rFOUND]
    aAdd($aR, $iPos)
  Next
  Return SetError(0, $mRegEx[$rOFFSET], $aR)
EndFunc   ;==>aRegExLocate

Func aRegExFindAll(Const $sInput, Const $sPattern, $iStart = 1, $sSep = "|")
  ; Features:
  ; [1] Where aRegEx finds one match, aRegExFindAll finds all the matches.
  ; [2] Unlike aRegEx, the total number of matches is returned in the zero
  ;     element of the main array.
  ; [3] If there is a match, by default element [1] contains a string with
  ;     the captured field(s).
  ; [4] If there is more than one capture per match, the values in the string
  ;     are separated by a vertical bar character '|'. Optionally choose
  ;     an other separator string.
  ; [5] If an empty string is chosen as separator, the function returns
  ;     array-in-array instead of string-in-array. Each row starts with the
  ;     zero element which contains the count of capture fields for that match.
  ; [6] For next matches, the main array has additional rows.
  ; [7] The string-in-array as well as the aDroste array-in-array technique
  ;     allow for variable numbers of captured fields.
  ; [8] There is no programmed maximum of matches that the function can return.
  ;
  ; Parameters:.. + $sInput..... String that contains the source of the data.
  ;               + $sPattern... The RegEx search criteria.
  ;               + $iStart..... [Optional] location in $sInput where the
  ;                              search starts. Default is 1.
  ;               + $sSep....... [Optional] denotes the format of the returned
  ;                              array. If $sSep is a character or non-empty
  ;                              string, each row in the array is a string
  ;                              where values are separated by $sSep.
  ;                              "|"  is the default, the vertical bar char.
  ;                              If $sSep is an empty string, a two-dimensional
  ;                              array is returned.
  ; Returns:......+ An array of matches, see features. The element with index
  ;                 zero contains the number of matches found.
  ; AutoIT bonus: + Macro @extended points to the location after the last match,
  ;                 so can be used to detect a footer in the input source.

  Local $Cap, $aR = aNewArray(), $mRegEx = mRegExdotInit($iStart)
  Local $I, $gRegEx = gRegEx($sInput, $sPattern, $iStart)
  For $I = 1 To $gRegEx[0]
    $mRegEx = $gRegEx[$I]
    $Cap = ($sSep = "") ? $mRegEx[$rCAPTURE] : sRecite($mRegEx[$rCAPTURE])
    aAdd($aR, $Cap)
  Next
  Return SetError(0, $mRegEx[$rOFFSET], $aR)
EndFunc   ;==>aRegExFindAll

Func aRegExSplit($sInput, $sPattern, $iStart = 1)
  ; Features:
  ; [1] Splits the sInput string in parts. Resembles the StringSplit() function
  ;     but offers more choice in separators.
  ; [2] Another difference with StringSplit is that an empty sInput string
  ;     result in zero parts.
  ; [3] Even if zero parts are returned, the returned value is of type array.
  ;
  ; Parameters:.. + $sInput..... The string to search in.
  ;               + $sPattern... The RegEx search criteria.
  ;               + $iStart..... [Optional] up to this location, the $sInput is
  ;                              ignored. Default is no ignoring.
  ; Returns:..... An array with the parts. The count property of the array
  ;               contains the number of parts.

  Local $aSplit = aNewArray(), $mRegEx = mRegExdotInit($iStart)
  Local $I, $gRegEx = gRegEx($sInput, $sPattern, $iStart)
  For $I = 1 To $gRegEx[0]
    $mRegEx = $gRegEx[$I]
    aAdd($aSplit, mRegExdotHeader($sInput, $mRegEx))
  Next
  If $mRegEx[$rOFFSET] <= StringLen($sInput) Then aAdd($aSplit, mRegExdotFooter($sInput, $mRegEx))
  Return $aSplit
EndFunc   ;==>aRegExSplit

Func sRegExReplace($sInput, $sPattern, $sScript, $iStart = 1)
  ; Features:
  ; [1] Finds all matches and replaces these by the outcome of a script.
  ;     With capture fields in the pattern and placeholders in the script,
  ;     parts of the match may return in the replacements.
  ; [2] Direct access in the replace script to the full collection of AutoIT
  ;     build-in and user defined functions. No prior declaration of these
  ;     functions is necessary.
  ; [4] This is the first RegExReplace implementation with this functionality
  ;     and is published before similar features exist in Java, Perl, Pcre,
  ;     Net, Python, Php and Ruby.
  ; [5] If for a back-reference in the script is no corresponding capture
  ;     field in the pattern, an informative message is issued at the console
  ;     and the program halts.
  ;
  ; Parameters:.. + $sInput..... The string to change
  ;               + $sPattern... The RegEx search string
  ;               + $sScript.... Describes the replacement. May include
  ;                              back-references and call-outs.
  ;               + $iStart..... (Optional) location in $sInput where search
  ;                              starts.
  ; Returns:..... + String with all the replacements made.
  ; AutoIT bonus: + Macro @extended contains the total number of matches.
  ;
  ; Explanation:
  ;     Use of this function needs some understanding of the 'capture fields'
  ;     mechanism. The pattern may contain one or more sub matches called
  ;     capture fields. These are parts of the full match. The function
  ;     replaces full matches with the outcome of the script, not just replace
  ;     a captured field. The contents of the captures - if any - can be
  ;     inserted anywhere in the script by means of placeholders, the so called
  ;     'back-references'.
  ;
  ; Back-references in the replace script:
  ;     For example: $1 references capture fields 1. An alternative syntax is: \1
  ;     In the output this placeholders is replaced with what is captured.
  ;
  ; Call-outs in the replace script:
  ;     The script may contain call-outs to any AutoIt or user- function. For
  ;     the syntax and a warning for the use in compiled programs, see the
  ;     sResolve() description.
  ;
  ; Note:
  ;     If one refrains from using the sCallOut() scripting feature, replace
  ;         line:   $sOut &= sCallOut($sBpart)
  ;         with:   $sOut &= $sBpart
  ;     If call-out is disabled, the backreference substitution mechanism still
  ;     works.

  ; Collect valid tokens from the replace script:
  Local $aaToken = aRegExFindAll($sScript, "(?<!\\)((?:\\|\$)(\d+))", 1, "")
  Local $aCap = aRegEx($sInput, $sPattern), $nCap = $aCap[0] ; ~
  If $nCap = 0 Then Return SetError(0, 0, $sInput)

  For $aToken In msBrace($aaToken) ; Check validity of the numbers:
    If $aToken[2] <= $nCap Then ContinueLoop
    Local $ePos = StringInStr($sScript, $aToken[1])
    cc('sReplace pattern error: ' & $sScript & @CRLF & Spaces($ePos + 24) & "^")
    Exit
  Next

  ; Capture headers, fragments and one footer:
  Local $sOut = "", $nFound = 0, $mRegEx = mRegExdotInit($iStart)
  Local $I, $gRegEx = gRegEx($sInput, $sPattern, $iStart)
  For $I = 1 To $gRegEx[0]
    $mRegEx = $gRegEx[$I]
    $nFound += 1
    $sOut &= mRegExdotHeader($sInput, $mRegEx)
    Local $sBpart = $sScript
    For $aToken In msBrace($aaToken)
      $sBpart = StringReplace($sBpart, $aToken[1], ($mRegEx[$rCAPTURE])[$aToken[2]])
    Next
    $sOut &= sCallOut($sBpart) ; Replace scripted functions with their return values
  Next
  $sOut &= mRegExdotFooter($sInput, $mRegEx)
  Return SetError(0, $nFound, $sOut)
EndFunc   ;==>sRegExReplace

#EndRegion ; Regulated expressions interface with mapped properties ==============

#Region ; AutoIT scripting probe ==============================================

Func sCallOut($sScript)
  ; Parameters:.. + $sScript.... The script to execute.
  ; Returns:..... + The resolved script with the results of the call-out
  ;                 function(s) filled in.
  ;
  ; Call-outs:
  ;     $sScript may contain call-outs to any AutoIt function. The syntax is:
  ;
  ;          <®fname®$1®bold®>
  ;
  ;     Where fname is the function name. This will call the AutoIT function
  ;     with "$1" as first parameter and "bold" as second parameter. Works
  ;     with zero to three parameters, separated from each other by the ®
  ;     character. Upon execution, the calling syntax including the less-than
  ;     and greater-than hooks are replaced by the return value. If the
  ;     return value is an array, the elements are converted to string
  ;     elements, delimited by the vertical bar character "|".
  ;
  ; Warning:
  ;     The use of Auto-IT function Call() needs care if the program is
  ;     compiled, because the functions that will be used in the script are not
  ;     known beforehand. Upon execution of the compiled program, call() may
  ;     not correctly link to the desired function. The collection of functions
  ;     to be available for the end-user [1] need to be present in the
  ;     source-code and [2] should not be stripped by the Au3Stripper.
  ;     Besides: [3] the Au3Stripper should not abbreviate function names.
  ;     Use the #Au3Stripper.ignore.funcs= directive to retain the full
  ;     names in the compiled code. For example, in the AutoIT3Wrapper heading
  ;     of this source file, the functions StringTitleCase() and StringFloat()
  ;     are exempted from stripping.

  Local $sBpart, $sMess = "AutoIT sRegExReplace script syntax report"
  While True
    Local $iFound1 = StringInStr($sScript, "<®")
    Local $iFound2 = $iFound1 ? StringInStr($sScript, "®>", 0, 1, $iFound1) : 0
    If $iFound2 = 0 Then ExitLoop
    Local $sInput = StringMid($sScript, $iFound1 + 2, $iFound2 - $iFound1 - 2)
    Local $aP = StringSplit($sInput, "®")
    If $aP[0] = 1 Then
      $sBpart = Call($aP[1])
    ElseIf $aP[0] = 2 Then
      $sBpart = Call($aP[1], $aP[2])
    ElseIf $aP[0] = 3 Then
      $sBpart = Call($aP[1], $aP[2], $aP[3])
    ElseIf $aP[0] = 4 Then
      $sBpart = Call($aP[1], $aP[2], $aP[3], $aP[4])
    EndIf
    If @error Or $aP[0] = 0 Or $aP[0] > 4 Then
      MsgBox(64, $sMess, "Call " & StringReplace($sInput, "®", " ") & @LF & @LF & "@error= 0xDEAD")
      ExitLoop
    EndIf
    If IsArray($sBpart) Then $sBpart = sRecite($sBpart)
    $sScript = StringLeft($sScript, $iFound1 - 1) & $sBpart & StringMid($sScript, $iFound2 + 2)
  WEnd
  Return $sScript
EndFunc   ;==>sCallOut

#EndRegion ; AutoIT scripting probe ==============================================

#Region ; String support =======================================================

Func cc(Const $sInput)
  ConsoleWrite($sInput & @CRLF)
EndFunc   ;==>cc

Func Crep(Const $Char, Const $nTimes = 10)
  Local $C = $Char, $N = $nTimes, $sResult = ""
  While $N > 1
    If BitAND($N, 1) Then $sResult &= $C
    $C &= $C
    $N = BitShift($N, 1)
  WEnd
  Return $N ? $C & $sResult : ""
EndFunc   ;==>Crep

Func Spaces(Const $N)
  Return Crep(' ', $N)
EndFunc   ;==>Spaces

Func sCell(Const $sInput, $nWidth)
  ; Helper routine for sJust()
  Local $sIn = StringStripWS($sInput, 3)
  $sIn = StringLeft($sIn, $nWidth)
  Local $Len = StringLen($sIn)
  Return ($nWidth > 0) ? Spaces($nWidth - $Len) & $sIn : $sIn & Spaces(Abs($nWidth) - $Len)
EndFunc   ;==>sCell
; MsgBox(0, 'TestsCell', sCell("United States", 15))

Func sJust(Const $Input, $nWidth = 15)
  ; Bring string- and number values in text format
  ; Input can be single value, 1dim array or 2dim array.
  ; Negative width denotes left alignment. Default is right alignment.
  Local $sOut, $nType = IsArray($Input) ; 1dim
  If $nType Then ; Check 2dim
    For $Elem In msBrace($Input)
      If IsArray($Elem) Then
        $nType += 1 ; 2dim
        ExitLoop
      EndIf
    Next
  EndIf
  Switch $nType
    Case 0 ; Single value
      $sOut = sCell($Input, $nWidth)
    Case 1 ; 1dim array
      For $sElem In msBrace($Input)
        $sOut &= sCell($sElem, $nWidth)
      Next
      $sOut &= @CRLF
    Case 2 ; 2dim array
      For $Line In msBrace($Input)
        For $sElem In msBrace($Line)
          $sOut &= sCell($sElem, $nWidth)
        Next
        $sOut = StringStripWS($sOut, 2) & @CRLF
      Next
  EndSwitch
  Return $sOut
EndFunc   ;==>sJust

Func DisplayBox($sTitle, $sContent = "")
  GUICreate($sTitle, 600, 406) ;
  GUICtrlCreateEdit($sContent, 0, 0, 600, 406, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $WS_HSCROLL, $WS_VSCROLL))
  GUICtrlSetFont(-1, 10, 400, 0, "Courier New")
  GUICtrlSetBkColor(-1, 0xFFFF80)
  GUICtrlSetState(-1, 256)
  GUISetState(@SW_SHOW)
  While True
    Local $nMsg = GUIGetMsg()
    Switch $nMsg
      Case $GUI_EVENT_CLOSE
        Exit
    EndSwitch
  WEnd
EndFunc   ;==>DisplayBox
; DisplayBox("Title", "Content")

#EndRegion ; String support =======================================================

#Region ; Number support =======================================================

Func StringFloat(Const $Number, Const $N)
  ; ReFormat float : # digits after comma
  If $N = 0 Then Return Round($Number)
  Local $Sign = $Number < 0, $Whole = Floor(Abs($Number))
  Local $Fract = Round((Abs($Number) - $Whole) * 10 ^ $N)
  If $Fract = 10 ^ $N Then
    $Whole += 1
    $Fract = StringTrimLeft($Fract, 1)
  EndIf
  If $Sign Then $Whole = -1 * $Whole
  Return $Whole & '.' & StringLeft($Fract & Crep('0', $N), $N)
EndFunc   ;==>StringFloat

#EndRegion ; Number support =======================================================

#Region ; RegEx examples =======================================================

Func sTitleCase2($sInput)
  ; A common practice is not to use 'regex' in the name of a function if it
  ; doesn't has a pattern as parameter, even when the working relies on
  ; RegEx technique.
  Return sRegExReplace($sInput, "(?:\A|\h+)\p{Ll}", "<®StringUpper®$1®>")
EndFunc   ;==>sTitleCase2

Func sTitleCase($sInput)
  Local $sOut = "", $mRegEx = mRegExdotInit(1)
  Local $I, $gRegEx = gRegEx($sInput, "(?:\A|\h+)\p{Ll}", 1)
  For $I = 1 To $gRegEx[0]
    $mRegEx = $gRegEx[$I]
    $sOut &= mRegExdotHeader($sInput, $mRegEx) & StringUpper(($mRegEx[$rCAPTURE])[1])
  Next
  $sOut &= mRegExdotFooter($sInput, $mRegEx)
  Return $sOut
EndFunc   ;==>sTitleCase

Func TestsTitleCase()
  ; A function _StringTitleCase is delivered with AutoIT in the Include folder
  ;  see: String.au3. Here are two alternatives.
  Local $S = "Although it started life as a simple automation tool, AutoIt now has functions " & _
      "and features that allow it to be used as a general purpose scripting language."
  MsgBox(0, "TitleCase", sTitleCase($S))
  MsgBox(0, "TitleCase2", sTitleCase2($S))
EndFunc   ;==>TestsTitleCase
; TestStringTitleCase()

Func CsvReader()
  ; Census bureau's offer tables on the Internet. This is how you can get these
  ; into your system. A table about renewable energy is supplied together with
  ; this source on GitHub.
  Local $sFile = FileOpenDialog("CsvReader", @MyDocumentsDir, "Table (*.csv;*.txt)", 3)
  If @error Then Exit
  Local $sInput = FileRead($sFile), $aOut = aNewArray(0), $mRegEx = mRegExdotInit(1)
  Local $I, $gRegEx = gRegEx($sInput, "((?:[\w\- ]+;)+[\w\- ]+)\r*\n", 1)
  Local $sHeader = mRegExdotHeader($sInput, Lif($gRegEx[0], $gRegEx[1], $mRegEx))
  For $I = 1 To $gRegEx[0]
    $mRegEx = $gRegEx[$I]
    aAdd($aOut, aRegExSplit(sRecite($mRegEx[$rCAPTURE]), ";"))
  Next
  DisplayBox("CsvReader", $sHeader & sJust($aOut) & mRegExdotFooter($sInput, $mRegEx))
EndFunc   ;==>CsvReader
; CsvReader()

Func NeedleInHaystack()
  ; fRits said me, man you're so busy with the programs, don't forget to add
  ; sample patterns. The users want examples. Adapt the collection to your needs.
  ; This program can scan folders such as Program files\AutoIT3\Examples.
  ; A file 'NeedleInHaystackMaterials.txt' is delivered together with this
  ; source code on GitHub.
  aD("Email address", "\b[\w+\.]{5,}@[\w]{3,}.\w{2,5}\b")
  aD("Dutch phone number", "\b\(*\d{3,4}(?:\h\-\h|\-|\h)\d{6,7}\b")
  aD("American phone number", "\s\([2-9][0-8][0-9]\)[2-9][0-9]{2}[\-\.][0-9]{4}\s")
  aD("International phone number", "\b\+[0-9]{1,3}\h[0-9\-]{4,14}\b")
  aD("ISBN", "\bISBN:\h\d\d\d\-\d-\d\d\d\-\d\d\d\d\d\-\d\b")
  aD("Canadian postal code", "\b[A-VXY][0-9][A-Z]\h[0-9][A-Z][0-9]\b")
  aD("UK postal code", "\b[A-Z]{1,2}[0-9][A-Z]\h[0-9][A-Z]{2}\b")
  aD("URL", "\b(?:https:\/\/|ftp:\/\/|www\/)\w+\.\w{2,5}\b")
  Local $sRoot, $aFiles = aSelectFiles($sRoot, 'Burst', 'txt|au3')
  If @error Then Exit
  Local $sOut = $sRoot & @CRLF, $aaPattern = aD(), $RootOffset = StringLen($sRoot) + 1, $N = 0
  For $File In msBrace($aFiles)
    Local $H = FileOpen($File), $sString = FileRead($H), $Lfirst = True
    FileClose($H)
    For $aPattern In msBrace($aaPattern)
      Local $aFound = aRegExFindAll($sString, $aPattern[2])
      If $aFound[0] Then
        If $Lfirst Then
          $sOut &= Spaces(2) & StringMid($File, $RootOffset) & @CRLF
          $Lfirst = False
        EndIf
        $sOut &= Spaces(4) & $aPattern[1] & ":" & @CRLF
        $N += $aFound[0]
      EndIf
      For $sFound In msBrace($aFound)
        $sOut &= Spaces(6) & $sFound & @CRLF
      Next
    Next
  Next
  DisplayBox("Needle in Haystack", $sOut & $N & ' found.' & @CRLF)
EndFunc   ;==>NeedleInHaystack
; NeedleInHaystack()

#EndRegion ; RegEx examples =======================================================
