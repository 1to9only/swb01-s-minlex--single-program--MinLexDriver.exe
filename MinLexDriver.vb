Imports System
Imports System.Threading.Tasks
' Copyright 2021, 2022 Shelby W. Blythe (C) SWB01X@gmail.com
' ************************************************************************************
' *         This program is free software: you can redistribute it and/or modify     *
' *     it under the terms of the GNU General Public License as published by the     *
' *     Free Software Foundation, Version 3 and any later version.                   *
' *                                                                                  *
' *     This program Is distributed in the hope that it will be useful,              *
' *     but WITHOUT ANY WARRANTY; without even the implied warranty of               *
' *     MERCHANTABILITY Or FITNESS FOR A PARTICULAR PURPOSE.  See the                *
' *     GNU General Public License for more details.                                 *
' *                                                                                  *
' *     You should have received a copy of the GNU General Public License            *
' *     along with this program.  If Not, see < https: //www.gnu.org/licenses/>.     *
' ************************************************************************************
' ************************************************************************************
' * With modifications by 1to9only                                                   *
' *                                                                                  *
' * This is shorter and much modified version of the original source code.           *
' * It supports old process mode 2 only, no need to specify the option.              *
' * Most of unused code, e.g. code for process mode 3, has been removed.             *
' *                                                                                  *
' * The minlexing code (MinLex9X9SR1) has seen some cosmetic changes.                *
' * The program usage has also been changed.                                         *
' *                                                                                  *
' * Usage: MinLexDriver.exe inputfile outputfile                                     *
' * e.g.                                                                             *
' * MinLexDriver.exe
' * MinLexDriver.exe ......789......1.....2.3...2......65.9..71.............7..9.8..............5...4.
' * MinLexDriver.exe 17puz49158.txt output.txt
' *                                                                                  *
' * For the input file, the restriction of 83-chars for each line remains.           *
' *                                                                                  *
' ************************************************************************************
Module Program
    Public MinLexDriverVersion As String = "VB-2025_09_09"

    Public InputBuffer1Chr() As Char = New Char(0) {}
    Public Buffer1Size As Integer
    Public MinLexBuffer1Offset As Integer
'   Public MinLex1ReturnCode As Integer         ' Values: 0: Success, 1: AscW conversion of conversion to integer failure, 2: input character < (0 or ".") or > 9.
'   Public ErrorInputRecord1 As String = Nothing

    Public InputBuffer2Chr() As Char = New Char(0) {}
    Public Buffer2Size As Integer
    Public MinLexBuffer2Offset As Integer
'   Public MinLex2ReturnCode As Integer
'   Public ErrorInputRecord2 As String = Nothing

    Public InputBuffer3Chr() As Char = New Char(0) {}
    Public Buffer3Size As Integer
    Public MinLexBuffer3Offset As Integer
'   Public MinLex3ReturnCode As Integer
'   Public ErrorInputRecord3 As String = Nothing

'   Public MinLexBufferStr() As String = New String(0) {}  ' Results are returned in this string array if a sort is required (process modes "1", "3", "4" or "6").
    Public MinLexBufferChr() As Char = New Char(0) {}      ' Results are returned in this character array if unsorted (process modes "2" or "5").
    Public MultiThreadTasksAreCompleteSw As Boolean
'   Public MinLexReturnCode As Integer
'   Public ErrorInputRecord As String = Nothing
'   Public ProcessMode As String = "2"
'   Public PatternModeSw As Boolean = False

    Sub Main(args As String())
'       Dim i, aix, bix, cix, lastix, resultsbufferix As Integer
        Dim InputFileFullPath As String = Nothing
        Dim OutputFileFullPath As String = Nothing
'       Dim MasterFileFullPath As String = Nothing
'       Dim MasterRecordStr As String
'       Dim LastMasterRecordStr As String
        Dim InputFileSize As Integer
        Dim InputPuzzleCount As Integer
'       Dim InputFolderPath As String = Nothing
'       Dim HoldStatusReport_Text As String
'       Dim MinLexDriverStartDateTime As DateTime
'       Dim CurrentDateTime As DateTime
'       Dim RunTimeSpan As TimeSpan
'       Dim ReadResults1 As Integer
'       Dim ReadResults2 As Integer
'       Dim ReadResults3 As Integer
'       Dim MasterFileSize As Integer
        Dim MinLexBuffer1StrIx As Integer
        Dim MinLexBuffer2StrIx As Integer
        Dim MinLexBuffer3StrIx As Integer
'       Dim DuplicateCount As Integer
'       Dim MasterFileCount As Integer
'       Dim MasterFileDuplicateCount As Integer
'       Dim Merge2EndSw As Boolean
'       Dim Merge3EndSw As Boolean
'       Dim MergeWithMasterEndSw As Boolean
        Dim InputRecordCountOneThirdPoint As Integer
        Dim MultiThreadFileSizeThreshold As Integer = 830000   ' 83 X 10,000   ' Input files with fewer than 10,000 puzzles are processed in a single buffer.
'       Dim ResultsBufferStr As String()                       ' Used for ProcessMode = "3" - Merge results with Master file.
'       Dim PromptModeSw As Boolean = False
'       Dim ValidInputArgumentsSw As Boolean = True
'       Dim OriginalProcessMode As String

        If args.Length <> 1 And args.Length <> 2 Then
            Console.WriteLine(MinLexDriverVersion & ", Usage: MinLexDriver.exe inputfile outputfile")
            Exit Sub
        End If
        If args.Length = 1 And Len(args(0)) >= 81 Then
            ReDim InputBuffer1Chr(83)
            Buffer1Size = 83
            MinLexBuffer1Offset = 0
            ReDim MinLexBufferChr(83)
            Array.ConstrainedCopy(args(0).ToCharArray, 0, InputBuffer1Chr, 0, 81)
            Dim MinLex1 As New Class1
'           MinLex1ReturnCode =
            MinLex1.MinLex9X9SR1(InputBuffer1Chr, Buffer1Size, MinLexBufferChr, MinLexBuffer1Offset)
            Console.WriteLine(MinLexBufferChr)
            Exit Sub
        End If
        If args.Length = 2 Then
            InputFileFullPath = args(0)
            OutputFileFullPath = args(1)
            Dim streamReaderTest As System.IO.FileStream
            Dim fileinfo As New IO.FileInfo(InputFileFullPath)
            Try
                streamReaderTest = System.IO.File.Open(InputFileFullPath, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            Catch ex As Exception
                Console.WriteLine("Error: Input file: " & ex.Message)
                Exit Sub
            End Try
            streamReaderTest.Close()
            InputFileSize = CInt(fileinfo.Length)
            If InputFileSize Mod 83 <> 0 Then
                Console.WriteLine("Error: Input file: Length error: Each line = 81-chars sudoku + CRLF.")
                Exit Sub
            End If
            InputPuzzleCount = InputFileSize \ 83
            If InputFileSize < MultiThreadFileSizeThreshold Then
                Buffer3Size = InputFileSize
                MinLexBuffer3Offset = 0
                ReDim InputBuffer3Chr(Buffer3Size)
                MinLexBuffer1StrIx = -1
                MinLexBuffer3StrIx = InputPuzzleCount - 1
            Else
                InputRecordCountOneThirdPoint = InputPuzzleCount \ 3

                Buffer1Size = InputRecordCountOneThirdPoint * 83
                MinLexBuffer1Offset = 0
                ReDim InputBuffer1Chr(Buffer1Size)
                MinLexBuffer1StrIx = InputRecordCountOneThirdPoint - 1

                Buffer2Size = Buffer1Size
                MinLexBuffer2Offset = InputRecordCountOneThirdPoint
                ReDim InputBuffer2Chr(Buffer2Size)
                MinLexBuffer2StrIx = 2 * InputRecordCountOneThirdPoint - 1

                Buffer3Size = InputFileSize - Buffer1Size - Buffer2Size
                MinLexBuffer3Offset = 2 * InputRecordCountOneThirdPoint
                ReDim InputBuffer3Chr(Buffer3Size)
                MinLexBuffer3StrIx = InputPuzzleCount - 1
            End If
            ReDim MinLexBufferChr(InputFileSize)
'           ReDim MinLexBufferStr(InputPuzzleCount)
            Dim streamReader As New System.IO.StreamReader(InputFileFullPath, System.Text.Encoding.ASCII, False, MultiThreadFileSizeThreshold)
            If streamReader.Peek() = -1 Then
                streamReader.Close()
                Console.WriteLine("Error: Input File: Empty File.")
                Exit Sub
            End If
            Dim streamWriterTest As System.IO.FileStream
            Try
                streamWriterTest = System.IO.File.Create(OutputFileFullPath)
            Catch ex As Exception
                streamReader.Close()
                Console.WriteLine("Error: Output file: " & ex.Message)
                Exit Sub
            End Try
            streamWriterTest.Close()
            If InputFileSize >= MultiThreadFileSizeThreshold Then
                MultiThreadTasksAreCompleteSw = False
                streamReader.ReadBlock(InputBuffer1Chr, 0, Buffer1Size)
                streamReader.ReadBlock(InputBuffer2Chr, 0, Buffer2Size)
                Dim MinLexAsync As New Task(AddressOf MinLexAsyncTasks)
                MinLexAsync.Start()
            Else
                MultiThreadTasksAreCompleteSw = True
            End If
            streamReader.ReadBlock(InputBuffer3Chr, 0, Buffer3Size)
            Dim MinLex3 As New Class1
'           MinLex3ReturnCode =
            MinLex3.MinLex9X9SR1(InputBuffer3Chr, Buffer3Size, MinLexBufferChr, MinLexBuffer3Offset)
            Do Until MultiThreadTasksAreCompleteSw
                Threading.Thread.Sleep(10)
            Loop
            streamReader.Close()
            Dim streamWriter As New System.IO.StreamWriter(OutputFileFullPath, False, System.Text.Encoding.ASCII, MultiThreadFileSizeThreshold)
'           If MinLex1ReturnCode > 0 Then
'               MinLexReturnCode = MinLex1ReturnCode
'               ErrorInputRecord = ErrorInputRecord1
'           ElseIf MinLex2ReturnCode > 0 Then
'               MinLexReturnCode = MinLex2ReturnCode
'               ErrorInputRecord = ErrorInputRecord2
'           ElseIf MinLex3ReturnCode > 0 Then
'               MinLexReturnCode = MinLex3ReturnCode
'               ErrorInputRecord = ErrorInputRecord3
'           End If
            streamWriter.Write(MinLexBufferChr, 0, InputFileSize)
            streamWriter.Close()
        End If
    End Sub ' Main

    Async Sub MinLexAsyncTasks()
        Dim MinLex1 As New Class1
        Dim MinLex1Awaiter As Task = Task.Run(
            Sub()
'               MinLex1ReturnCode =
                MinLex1.MinLex9X9SR1(
                    InputBuffer1Chr,
                    Buffer1Size,
                    MinLexBufferChr,
                    MinLexBuffer1Offset
                    )
'                   MinLexBufferStr,
'                   ProcessMode,
'                   PatternModeSw
'                   ErrorInputRecord1
            End Sub)
        Dim MinLex2 As New Class1
        Dim MinLex2Awaiter As Task = Task.Run(
            Sub()
'               MinLex2ReturnCode =
                MinLex2.MinLex9X9SR1(
                    InputBuffer2Chr,
                    Buffer2Size,
                    MinLexBufferChr,
                    MinLexBuffer2Offset
                    )
'                   MinLexBufferStr,
'                   ProcessMode,
'                   PatternModeSw
'                   ErrorInputRecord2
            End Sub)
        Await MinLex1Awaiter
        Await MinLex2Awaiter
        MultiThreadTasksAreCompleteSw = True
    End Sub

End Module

Public Class Class1
    Public Function MinLex9X9SR1(InputGridBufferChr As Char(),
                                 InputBufferSize As Integer,
                                 MinLexBufferChr As Char(),
                                 ResultsBufferOffset As Integer) As Integer
'                                MinlexBufferStr As String(),
'                                ProcessMode As String,
'                                PatternModeSw As Boolean) As Integer              ' True = input processed as a pattern; False process as grid or sub-grid.
'                          ByRef ErrorInputRecordStr As String) As Integer
'
        Dim a, b, i, j, k, inputgridstart, m, n, z, zstart, candidate, hold, inputgridix As Integer
        Dim CarriageReturnChr As Char = ChrW(AscW("0"c) - 35)
        Dim LineFeedChr As Char = ChrW(AscW("0"c) - 38)
        Dim ResultsBufferCharacterOffset As Integer = ResultsBufferOffset * 83
        Dim NextChr As Char
        Dim InputLineCount As Integer
        Dim CluesCount As Integer
        Dim trackercolumnpermutationix As Integer
        Dim row4stackpermutationix, row5stackpermutationix, row6stackpermutationix, row7stackpermutationix, row8stackpermutationix, row9stackpermutationix As Integer
        Dim row4columnpermutationix, row5columnpermutationix, row6columnpermutationix, row7columnpermutationix, row8columnpermutationix, row9columnpermutationix As Integer
        Dim band2row4orderix, band2rows5and6orderix, band3row7orderix, band3rows8and9orderix As Integer
        Dim MinLexGridLocalReset As Integer() = New Integer(80) {10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}
        Dim MinLexGridLocal As Integer() = New Integer(80) {}
'       Dim MinLexGridLocalChr(80) As Char
        Dim CharPeriodTO9 As Char() = New Char(9) {"."c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
        Dim FirstNonZeroDigitPositionInRow As Integer() = New Integer(9) {}                                                        'NOTE: Zero & 1 elements not used.
        Dim RowPositionalWeight As Integer() = New Integer(9) {}                                                        'NOTE: Zero & 1 elements not used.
        Dim FirstNonZeroDigitRelabeled As Integer() = New Integer(9) {}                                                            'NOTE: Zero, 1, 2, 3 elements not used.
        Dim MinLexFirstNonZeroDigitPositionInRowReset As Integer() = New Integer(9) {0, 0, -1, -1, -1, -1, -1, -1, -1, -1}         'NOTE: Zero, 1, 2, 3 elements not used.
        Dim MinLexFirstNonZeroDigitPositionInRow As Integer() = New Integer(9) {0, 0, -1, -1, -1, -1, -1, -1, -1, -1}
'       Dim MinLexRowPositionalWeightInit As Integer() = New Integer(9) {0, 1024, 1024, 1024, 1024, 1024, 1024, 1024, 1024, 1024}
'       Dim MinLexRowPositionalWeight As Integer() = New Integer(9) {}
        Dim CandidateFirstRelabeledDigit As Integer
        Dim Row2TestFirstNonZeroDigitPositionInRow As Integer
        Dim Row3TestFirstNonZeroDigitPositionInRow As Integer
        Dim Row4TestFirstNonZeroDigitPositionInRow As Integer
        Dim Row5TestFirstNonZeroDigitPositionInRow As Integer
        Dim Row7TestFirstNonZeroDigitPositionInRow As Integer
        Dim Row8TestFirstNonZeroDigitPositionInRow As Integer
'       Dim Row2MinimumPositionalWeight As Integer
'       Dim Row3MinimumPositionalWeight As Integer
'       Dim Row4MinimumPositionalWeight As Integer
'       Dim Row5MinimumPositionalWeight As Integer
'       Dim Row7MinimumPositionalWeight As Integer
'       Dim Row8MinimumPositionalWeight As Integer
        Dim ColumnPermutationTrackerIx As Integer
        ' NOTE: This routine identifies the MinLex for Band1 examining all viable candidates. It "right-justifies"
        '       the Band1 candidate and stores the column/stack reconfigured puzzle off for later retrieval, if it yields the Band1 MilLex.
        '       For each "Band1 right-justified" puzzle, the following arrays store the column and stack permutations, and other details, necessary to
        '       reproduce the Band1 MinLex during subsequent processing for rows 4 - 9.
        '
        '       The column permutation trackers hold the column permutations that need to be applied to the right-justified Band1 candidates to
        '       match the MinLex of Band1 - using 0-8 notation. For example, if a tracker = "345021867", then
        '       to reproduce the Minlex, switch Columns 2 & 3, rotate columns 7, 8 & 9 and then switch Stacks 1 & 2.
        Dim ColumnPermutationTrackerArrayMax As Integer = 999                               ' 648, the maximum number of grid automorphisms, is probably sufficient for the number of permutations that need to be stored.
        Dim CandidateColumnPermutationTrackerStartIx As Integer
        Dim FirstColumnPermutationTrackerIsIdentitySw As Boolean
        Dim ColumnPermutationTrackerIsIdentitySw As Boolean() = New Boolean(999) {}         ' 1000 - 1     ' Identifies those Trackers that indicate no change (The "identity" permutation - "012345678")
        Dim ColumnPermutationTracker As Integer() = New Integer(8999) {}                    ' 1000 X 9 - 1
        Dim HoldDigitAlreadyHitSw As Boolean() = New Boolean(359) {}                        ' 36 X 10 - 1
        Dim HoldDigitsRelabelWrk As Integer() = New Integer(9999) {}                        ' 1000 X 10 - 1
        Dim HoldRelabelLastDigit As Integer() = New Integer(35) {}                          ' 36 - 1
        Dim HoldBand1CandidateJustifiedPuzzles As Integer() = New Integer(2915) {}          ' 36 X 81 - 1
        Dim ColumnTrackerInit As Integer() = New Integer(8) {0, 1, 2, 3, 4, 5, 6, 7, 8}
        Dim LocalColumnPermutationTracker As Integer() = New Integer(8) {}
        Dim MinLexRows As Integer() = New Integer(26) {}
        Dim MinLexRowHitSw As Boolean
        Dim ResetMinLexRowSw As Boolean

        Dim iForHit As Integer() = New Integer(8) {}
        Dim Row4TestCandidateRowIx As Integer
        Dim Row4TestCandidateRow As Integer() = New Integer(8) {}
        Dim Row4TestPositionalCandidateRow As Integer() = New Integer(9) {}
        Dim Row5TestCandidateRowIx As Integer
        Dim Row5TestCandidateRow As Integer() = New Integer(1) {}
        Dim Row7TestCandidateRowIx As Integer
        Dim Row7TestCandidateRow As Integer() = New Integer(2) {}
        Dim Row7TestPositionalCandidateRow As Integer() = New Integer(9) {}
        Dim Row8TestCandidateRowIx As Integer
        Dim Row8TestCandidateRow As Integer() = New Integer(1) {}

        Dim InputGrid As Integer() = New Integer(161) {}
        Dim TestBand1 As Integer() = New Integer(80) {}
        Dim TestBand2 As Integer() = New Integer(80) {}
        Dim TestBand3 As Integer() = New Integer(80) {}
        Dim MinLexCandidateSw As Boolean
        Dim CheckThisPassSw As Boolean
        Dim TestGridRelabeled As Integer() = New Integer(80) {}
        Dim DigitAlreadyHitSw As Boolean() = New Boolean(9) {}          ' Note: zero element not used.
        Dim Row3DigitAlreadyHitSw As Boolean() = New Boolean(9) {}      ' Note: zero element not used.
        Dim Row4DigitAlreadyHitSw As Boolean() = New Boolean(9) {}      ' Note: zero element not used.
        Dim Row5DigitAlreadyHitSw As Boolean() = New Boolean(9) {}      ' Note: zero element not used.
        Dim Row6DigitAlreadyHitSw As Boolean() = New Boolean(9) {}      ' Note: zero element not used.
        Dim Row7DigitAlreadyHitSw As Boolean() = New Boolean(9) {}      ' Note: zero element not used.
        Dim Row8DigitAlreadyHitSw As Boolean() = New Boolean(9) {}      ' Note: zero element not used.
        Dim Row9DigitAlreadyHitSw As Boolean() = New Boolean(9) {}      ' Note: zero element not used.

        Dim Row3RelabelLastDigit As Integer
        Dim Row4RelabelLastDigit As Integer
        Dim Row5RelabelLastDigit As Integer
        Dim Row6RelabelLastDigit As Integer
        Dim Row7RelabelLastDigit As Integer
        Dim Row8RelabelLastDigit As Integer
        Dim Row9RelabelLastDigit As Integer

        ' For DigitsRelabelWrk, the zero element relabels 0 to 0, the 10 element is used to assign the first digit in an empty row to position "10".
        Dim DigitsRelabelWrkInit As Integer() = New Integer(10) {0, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10}
        Dim Row3DigitsRelabelWrk As Integer() = New Integer(10) {}
        Dim Row4DigitsRelabelWrk As Integer() = New Integer(10) {}
        Dim Row5DigitsRelabelWrk As Integer() = New Integer(10) {}
        Dim Row6DigitsRelabelWrk As Integer() = New Integer(10) {}
        Dim Row7DigitsRelabelWrk As Integer() = New Integer(10) {}
        Dim Row8DigitsRelabelWrk As Integer() = New Integer(10) {}
        Dim Row9DigitsRelabelWrk As Integer() = New Integer(10) {}

        Dim TranspositionNeededSw As Boolean
        Dim OriginalMiniRowCount As Integer(,) = New Integer(17, 2) {}
        Dim JustifiedMiniRowCount As Integer(,) = New Integer(17, 2) {}
        Dim AugmentedMiniRowCount1 As Integer
        Dim AugmentedMiniRowCount2 As Integer
        Dim IntToBit1To9 As Integer() = New Integer(9) {0, 1, 2, 4, 8, 16, 32, 64, 128, 256}
        Dim DigitsInRowBit As Integer() = New Integer(17) {}
        Dim DigitsInMiniRowBit As Integer(,) = New Integer(17, 2) {}
        Dim CalcMiniRowCountCode As Integer() = New Integer(17) {}
        Dim CalcMiniRowCountCodeMinimum As Integer
        Dim ZeroRowsInBandsCount As Integer() = New Integer(5) {}
        Dim Row2CalcMiniRowCountCodeMinimum As Integer
        Dim RowWithCalcMiniRowCountCodeMinimumIx As Integer
        Dim RowWithCalcMiniRowCountCodeMinimum As Integer() = New Integer(17) {}
        Dim TwoRowCandidateIx As Integer
        Dim TwoRowCandidateRow1 As Integer() = New Integer(35) {}
        Dim TwoRowCandidateRow2 As Integer() = New Integer(35) {}
        Dim TwoRowCandidateRow3 As Integer() = New Integer(35) {}
        Dim TwoRowCandidateMiniRowCode As Integer() = New Integer(35) {}
        Dim TwoRowCandidateMiniRow1Count As Integer
        Dim TwoRowCandidateMiniRow2Count As Integer
        Dim CandidatesRepeatDigitsBetweenFirstAndSecondRowSw As Boolean
        Dim RowRepeatDigits As Integer
        Dim MiniRowRepeatDigits As Integer
        Dim TwoRowCandidateMiniRowCodeMinimum As Integer
        Dim TwoRowCandidateminirow1CodeMinimum As Integer
        Dim Step2Row1CandidateIx As Integer
        Dim Step2Row1Candidate As Integer() = New Integer(35) {}
        Dim Step2Row2Candidate As Integer() = New Integer(35) {}
        Dim Step2ColumnPermutationTrackerStartIx As Integer() = New Integer(35) {}
        Dim ColumnPermutationTrackerCount As Integer
        Dim Step2ColumnPermutationTrackerCount As Integer() = New Integer(35) {}
'       Dim Step2CandidateDigitAlreadyRelabeledSw As Boolean() = New Boolean(35) {}
        Dim MiniRowOrderTrackerInit As Integer(,) = New Integer(17, 2) {{0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}, {0, 1, 2}}
        Dim MiniRowOrderTracker As Integer(,) = New Integer(17, 2) {}
        Dim Band1CandidateIx As Integer
        Dim Band1CandidateRow1 As Integer() = New Integer(35) {}
        Dim Band1CandidateRow2 As Integer() = New Integer(35) {}
        Dim Band1CandidateRow3 As Integer() = New Integer(35) {}
'       Dim Band1CandidateDigitsRepeatInBandSw As Boolean() = New Boolean(35) {}
        Dim FirstNonZeroRowCandidateIx As Integer
        Dim Row1Candidate As Integer
        Dim Row2Candidate As Integer
        Dim Row3Candidate As Integer
        Dim Row1Candidate17 As Integer
        Dim Row2Candidate17 As Integer
        Dim Row3Candidate17 As Integer
        Dim Row1StackPermutationCode, Row2StackPermutationCode, Row2Or3StackPermutationCode, Row3StackPermutationCode, Row4StackPermutationCode, Row5StackPermutationCode, Row6StackPermutationCode, Row7StackPermutationCode, Row8StackPermutationCode, Row9StackPermutationCode As Integer

        ''                       Row 1 Candidate Rows =    1  2  3  4  5  6  7  8  9  10  11 12  13  14  15  16  17  18
        Dim Row2CandidateA As Integer() = New Integer(17) {1, 2, 0, 4, 5, 3, 7, 8, 6, 10, 11, 9, 13, 14, 12, 16, 17, 15}    ' Note: in 0 to 17 notation.
        Dim Row2CandidateB As Integer() = New Integer(17) {2, 0, 1, 5, 3, 4, 8, 6, 7, 11, 9, 10, 14, 12, 13, 17, 15, 16}

        Dim HoldPuzzle As Integer() = New Integer(80) {}

        '                                                                          0  1  2  3   4  5  6  7   8   9 10 11 12  13  14  15  16  17  18  19  20 21 22 23  24  25 26 27  28  29  30 31  32  33  34  35  36  37  38  39  40  41 42 43  44  45  46 47  48  49  50  51  52  53  54  55  56  57  58  59  60  61  62 63
        '                                                                             |  |  |      |  |  |          |  |              |                      |  |  |          |  |              |                                          |  |              |                                                              |
        Dim FirstNonZeroPositionInFirstNonZeroRow As Integer() = New Integer(63) {-1, 8, 7, 6, -1, 5, 5, 5, -1, -1, 4, 4, -1, -1, -1, 3, -1, -1, -1, -1, -1, 2, 2, 2, -1, -1, 2, 2, -1, -1, -1, 2, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 1, 1, -1, -1, -1, 1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 0}
        Dim FirstNonZeroPositionInRow1 As Integer
        Dim FirstNonZeroPositionInRow2 As Integer
        '                                                                       0  1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63
        '                                                                                      |              |              |                 |  |  |        |              |                                |  |           |                                               |
        Dim FirstNonZeroRowStackPermutationCode As Integer() = New Integer(63) {0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 3, 2, 2, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 2, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3}
        Dim StackPermutations As Integer() = New Integer(3) {0, 1, 1, 5}
        Dim PermutationStackX()() As Integer =
                                {({0}),
                                 ({0, 1}),
                                 ({0, 0}),
                                 ({0, 1, 0, 1, 0, 1})}
        Dim PermutationStackY()() As Integer =
                                {({0}),
                                 ({0, 2}),
                                 ({0, 1}),
                                 ({0, 2, 1, 2, 1, 2})}

        Dim Row4ColumnPermutationCode, Row5ColumnPermutationCode, Row6ColumnPermutationCode, Row7ColumnPermutationCode, Row8ColumnPermutationCode, Row9ColumnPermutationCode As Integer
        Dim ColumnPermutations As Integer() = New Integer(63) {0, 1, 1, 5, 1, 3, 3, 11, 1, 3, 3, 11, 5, 11, 11, 35, 1, 3, 3, 11, 3, 7, 7, 23, 3, 7, 7, 23, 11, 23, 23, 71, 1, 3, 3, 11, 3, 7, 7, 23, 3, 7, 7, 23, 11, 23, 23, 71, 5, 11, 11, 35, 11, 23, 23, 71, 11, 23, 23, 71, 35, 71, 71, 215}
        Dim PermutationColumnX()() As Integer =
                                 {({0}),
                                  ({0, 7}),
                                  ({0, 6}),
                                  ({0, 7, 6, 7, 6, 7}),
                                  ({0, 4}),
                                  ({0, 7, 4, 7}),
                                  ({0, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}),
                                  ({0, 3}),
                                  ({0, 7, 3, 7}),
                                  ({0, 6, 3, 6}),
                                  ({0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7}),
                                  ({0, 4, 3, 4, 3, 4}),
                                  ({0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7}),
                                  ({0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}),
                                  ({0, 1}),
                                  ({0, 7, 1, 7}),
                                  ({0, 6, 1, 6}),
                                  ({0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7}),
                                  ({0, 4, 1, 4}),
                                  ({0, 7, 4, 7, 1, 7, 4, 7}),
                                  ({0, 6, 4, 6, 1, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}),
                                  ({0, 3, 1, 3}),
                                  ({0, 7, 3, 7, 1, 7, 3, 7}),
                                  ({0, 6, 3, 6, 1, 6, 3, 6}),
                                  ({0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7}),
                                  ({0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4}),
                                  ({0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7}),
                                  ({0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}),
                                  ({0, 0}),
                                  ({0, 7, 0, 7}),
                                  ({0, 6, 0, 6}),
                                  ({0, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7}),
                                  ({0, 4, 0, 4}),
                                  ({0, 7, 4, 7, 0, 7, 4, 7}),
                                  ({0, 6, 4, 6, 0, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}),
                                  ({0, 3, 0, 3}),
                                  ({0, 7, 3, 7, 0, 7, 3, 7}),
                                  ({0, 6, 3, 6, 0, 6, 3, 6}),
                                  ({0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7}),
                                  ({0, 4, 3, 4, 3, 4, 0, 4, 3, 4, 3, 4}),
                                  ({0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7}),
                                  ({0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}),
                                  ({0, 1, 0, 1, 0, 1}),
                                  ({0, 7, 1, 7, 0, 7, 1, 7, 0, 7, 1, 7}),
                                  ({0, 6, 1, 6, 0, 6, 1, 6, 0, 6, 1, 6}),
                                  ({0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7}),
                                  ({0, 4, 1, 4, 0, 4, 1, 4, 0, 4, 1, 4}),
                                  ({0, 7, 4, 7, 1, 7, 4, 7, 0, 7, 4, 7, 1, 7, 4, 7, 0, 7, 4, 7, 1, 7, 4, 7}),
                                  ({0, 6, 4, 6, 1, 6, 4, 6, 0, 6, 4, 6, 1, 6, 4, 6, 0, 6, 4, 6, 1, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}),
                                  ({0, 3, 1, 3, 0, 3, 1, 3, 0, 3, 1, 3}),
                                  ({0, 7, 3, 7, 1, 7, 3, 7, 0, 7, 3, 7, 1, 7, 3, 7, 0, 7, 3, 7, 1, 7, 3, 7}),
                                  ({0, 6, 3, 6, 1, 6, 3, 6, 0, 6, 3, 6, 1, 6, 3, 6, 0, 6, 3, 6, 1, 6, 3, 6}),
                                  ({0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7}),
                                  ({0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4, 0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4, 0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4}),
                                  ({0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7}),
                                  ({0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6}),
                                  ({0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7})}
        Dim PermutationColumnY()() As Integer =
                                 {({0}),
                                  ({0, 8}),
                                  ({0, 7}),
                                  ({0, 8, 7, 8, 7, 8}),
                                  ({0, 5}),
                                  ({0, 8, 5, 8}),
                                  ({0, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}),
                                  ({0, 4}),
                                  ({0, 8, 4, 8}),
                                  ({0, 7, 4, 7}),
                                  ({0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8}),
                                  ({0, 5, 4, 5, 4, 5}),
                                  ({0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8}),
                                  ({0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}),
                                  ({0, 2}),
                                  ({0, 8, 2, 8}),
                                  ({0, 7, 2, 7}),
                                  ({0, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8}),
                                  ({0, 5, 2, 5}),
                                  ({0, 8, 5, 8, 2, 8, 5, 8}),
                                  ({0, 7, 5, 7, 2, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}),
                                  ({0, 4, 2, 4}),
                                  ({0, 8, 4, 8, 2, 8, 4, 8}),
                                  ({0, 7, 4, 7, 2, 7, 4, 7}),
                                  ({0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8}),
                                  ({0, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5}),
                                  ({0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8}),
                                  ({0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}),
                                  ({0, 1}),
                                  ({0, 8, 1, 8}),
                                  ({0, 7, 1, 7}),
                                  ({0, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8}),
                                  ({0, 5, 1, 5}),
                                  ({0, 8, 5, 8, 1, 8, 5, 8}),
                                  ({0, 7, 5, 7, 1, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}),
                                  ({0, 4, 1, 4}),
                                  ({0, 8, 4, 8, 1, 8, 4, 8}),
                                  ({0, 7, 4, 7, 1, 7, 4, 7}),
                                  ({0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8}),
                                  ({0, 5, 4, 5, 4, 5, 1, 5, 4, 5, 4, 5}),
                                  ({0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 1, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8}),
                                  ({0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 1, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}),
                                  ({0, 2, 1, 2, 1, 2}),
                                  ({0, 8, 2, 8, 1, 8, 2, 8, 1, 8, 2, 8}),
                                  ({0, 7, 2, 7, 1, 7, 2, 7, 1, 7, 2, 7}),
                                  ({0, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8}),
                                  ({0, 5, 2, 5, 1, 5, 2, 5, 1, 5, 2, 5}),
                                  ({0, 8, 5, 8, 2, 8, 5, 8, 1, 8, 5, 8, 2, 8, 5, 8, 1, 8, 5, 8, 2, 8, 5, 8}),
                                  ({0, 7, 5, 7, 2, 7, 5, 7, 1, 7, 5, 7, 2, 7, 5, 7, 1, 7, 5, 7, 2, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}),
                                  ({0, 4, 2, 4, 1, 4, 2, 4, 1, 4, 2, 4}),
                                  ({0, 8, 4, 8, 2, 8, 4, 8, 1, 8, 4, 8, 2, 8, 4, 8, 1, 8, 4, 8, 2, 8, 4, 8}),
                                  ({0, 7, 4, 7, 2, 7, 4, 7, 1, 7, 4, 7, 2, 7, 4, 7, 1, 7, 4, 7, 2, 7, 4, 7}),
                                  ({0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8}),
                                  ({0, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5, 1, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5, 1, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5}),
                                  ({0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 1, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 1, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8}),
                                  ({0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 1, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 1, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7}),
                                  ({0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8})}

        Dim FixedColumns As Integer() = New Integer(8) {}      ' Indicates if a column containes a non-zero digit above the current row. Elements 0 to 8 correspond to columns 1 to 9.
        Dim Row3FixedColumns As Integer() = New Integer(8) {}
        Dim Row4FixedColumns As Integer() = New Integer(8) {}
        Dim Row5FixedColumns As Integer() = New Integer(8) {}
        Dim Row6FixedColumns As Integer() = New Integer(8) {}
        Dim Row7FixedColumns As Integer() = New Integer(8) {}
        Dim Row8FixedColumns As Integer() = New Integer(8) {}

        Dim FixedColumnsSavedAsOfRow4Sw As Boolean
        Dim FixedColumnsSavedAsOfRow6Sw As Boolean
        Dim FixedColumnsSavedAsOfRow7Sw As Boolean

        Dim StillJustifyingSw As Boolean
        Dim StoppedJustifyingRow As Integer
        Dim HoldGrid As Integer() = New Integer(80) {}
        Dim HoldRow As Integer() = New Integer(8) {}
        Dim StartEqualCheck As Integer

        Dim stackx, stacky, switchx, switchy As Integer
        Dim row1start, row2start, row3start As Integer
        Dim Row1MiniRowCount As Integer() = New Integer(2) {}
'       Dim Row1MiniRowPermutationCode As Integer() = New Integer(2) {}
        Dim Row2MiniRowCount As Integer() = New Integer(2) {}
        Dim Row3MiniRowCount As Integer() = New Integer(2) {}
        Dim ColumnPermutationCode As Integer
        Dim Band1MiniRowColumnPermutationCode As Integer() = New Integer(2) {}
        Dim StackPermutationTrackerLocal As Integer() = New Integer(2) {}
        Dim StackPermutationTrackerCode As Integer
        Dim LocalBand1 As Integer() = New Integer(26) {}
        Dim LocalRows2and3 As Integer() = New Integer(18) {}
        Dim DigitsRelabelWrk As Integer() = New Integer(10) {}
        Dim RelabelLastDigit As Integer
        Dim TestRowsRelabeled As Integer() = New Integer(26) {}
        Dim stackpermutationix As Integer
        Dim columnpermutationix As Integer
        Dim FoundColumnPermutationTrackerCountMax As Integer
        ' End Variables For MinLexBand1 and MinlexRows2and3
        ' Variables for ProcessFullGrid
        Dim band1order, band1row1order As Integer
        Dim stackorder, columnorder As Integer
        Dim MinLexFullGridLocalReset As Integer() = New Integer(80) {1, 2, 3, 4, 5, 6, 7, 8, 9, 4, 5, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9}
        Dim TestGridBands As Integer() = New Integer(80) {}
        Dim TestGridRows As Integer() = New Integer(80) {}
        Dim TestGridCols As Integer() = New Integer(80) {}
        Dim TestGrid123 As Integer() = New Integer(80) {}
        Dim HoldCell1 As Integer
        Dim HoldCell2 As Integer
        Dim HoldCell3 As Integer
'       Dim DigitsOrderForMinLex As Integer() = New Integer(8) {}
'       Dim RowOrderForMinLex As Integer() = New Integer(8) {}
'       Dim ColumnOrderForMinLex As Integer() = New Integer(8) {}
        Dim LoopCount As Integer
        Dim MinLexType456Sw As Boolean
        Dim MinLexBandType456Sw As Boolean() = New Boolean(5) {}
        Dim MatchingDigit1 As Integer
        Dim MatchingDigit2 As Integer
        Dim MatchingDigit1Position As Integer
        Dim MatchingDigit2Position As Integer
        Dim MatchingDigitCount As Integer
        ' End Variables for ProcessFullGrid

        InputLineCount = InputBufferSize \ 83
        For inputgridix = 0 To InputLineCount - 1
            Array.Clear(OriginalMiniRowCount, 0, 54)
            Array.ConstrainedCopy(MinLexGridLocalReset, 0, MinLexGridLocal, 0, 81)
            TranspositionNeededSw = False
            Array.Clear(DigitsInMiniRowBit, 0, 54)
            ' Preliminary Step#1: Identify all row1 candidates based on minirow count pattern.
            ' Count non-zero digits in minirows for direct and transposed puzzle.
            CluesCount = 0
            Array.Clear(InputGrid, 0, 81)
            inputgridstart = inputgridix * 83
'           If PatternModeSw Then
'               For k = 0 To 80                       ' Convert input char puzzle to integer array. All characters other than zero and period changed to ones (1s).
'                   NextChr = InputGridBufferChr(inputgridstart + k)
'                   If Not (NextChr = "."c Or NextChr = "0"c) Then
'                       InputGrid(k) = 1                         ' Set to 1.
'                       CluesCount += 1
'                       i = k \ 9                                ' i = Row                 - 0 to 8 notation
'                       j = k Mod 9                              ' j = Column              - 0 to 8 notation
'                       m = j \ 3                                ' m = MiniRow or Stack    - 0 to 2 notation
'                       n = i \ 3                                ' n = MiniColumn or Band  - 0 to 2 notation
'                       OriginalMiniRowCount(i, m) += 1
'                       OriginalMiniRowCount(j + 9, n) += 1
'                   End If
'               Next k
'           Else              ' For sub-grid and full-grid processing.
                For k = 0 To 80                       ' Convert input char puzzle to integer array. Each non-zero (or non-period) character must be in range of 1 to 9.
                    NextChr = InputGridBufferChr(inputgridstart + k)
                    If NextChr >= "1"c And NextChr <= "9"c Then
'                       If NextChr > "0"c Then
                            InputGrid(k) = AscW(NextChr) - 48      ' Convert Char to Integer.
                            CluesCount += 1
                            i = k \ 9                                ' i = Row                 - 0 to 8 notation
                            j = k Mod 9                              ' j = Column              - 0 to 8 notation
                            m = j \ 3                                ' m = MiniRow or Stack    - 0 to 2 notation
                            n = i \ 3                                ' n = MiniColumn or Band  - 0 to 2 notation
                            OriginalMiniRowCount(i, m) += 1
                            OriginalMiniRowCount(j + 9, n) += 1
                            DigitsInMiniRowBit(i, m) = DigitsInMiniRowBit(i, m) Or IntToBit1To9(InputGrid(k))          ' Turn on digit's bit in direct row's minirow.
                            DigitsInMiniRowBit(j + 9, n) = DigitsInMiniRowBit(j + 9, n) Or IntToBit1To9(InputGrid(k))  ' Turn on digit's bit in transposed row's minirow.
'                       End If
'                   Else
'                       ErrorInputRecordStr = Nothing
'                       For j = 0 To 80
'                           ErrorInputRecordStr += InputGridBufferChr(inputgridstart + j)
'                       Next j
'                       Return 1   ' Error 1 , Invalid character in input record. Must contain only digits 0 - 9 (with ""."" treated as 0).
                    End If
                Next k
'           End If
'               ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! sub grid (puzzles) Start !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If CluesCount < 81 Then   ' Process sub grid (puzzles)
                ' "Rightjustify" MiniRowCounts for rows of direct and transposed grids.
                '  NOTE: Isolated performance testing indicates that, for array size less than 20, a for loop is faster than Array.Clear() and Array.ConstrainedCopy().
                '        So they have been replaces by equivalent for loops for low array sizes. The origianl subroutine calls remain as comments.
                'Array.Clear(ZeroRowsInBandsCount, 0, 6)
                For z = 0 To 5 : ZeroRowsInBandsCount(z) = 0 : Next z
                'Array.ConstrainedCopy(MinLexFirstNonZeroDigitPositionInRowReset, 4, MinLexFirstNonZeroDigitPositionInRow, 4, 6)
                For z = 4 To 9 : MinLexFirstNonZeroDigitPositionInRow(z) = MinLexFirstNonZeroDigitPositionInRowReset(z) : Next z
                Array.ConstrainedCopy(OriginalMiniRowCount, 0, JustifiedMiniRowCount, 0, 54)
                Array.ConstrainedCopy(MiniRowOrderTrackerInit, 0, MiniRowOrderTracker, 0, 54)
                CalcMiniRowCountCodeMinimum = 99
                For i = 0 To 17     ' Right juastify direct and transposed band counts for eack row - high count to the right, e.g. 1,3,0 becomes 0,1,3.
                    DigitsInRowBit(i) = DigitsInMiniRowBit(i, 0) Or DigitsInMiniRowBit(i, 1) Or DigitsInMiniRowBit(i, 2)
                    If DigitsInRowBit(i) = 0 Then
                        CalcMiniRowCountCode(i) = 0
                        ZeroRowsInBandsCount(i \ 3) += 1
'                       If ZeroRowsInBandsCount(i \ 3) > 1 Then
'                           ErrorInputRecordStr = Nothing
'                           For j = 0 To 80
'                               ErrorInputRecordStr += CStr(InputGridBufferChr(inputgridstart + j))
'                           Next j
'                           Return 2   ' Error 2, more than one all zeros Row or Column in a Band or Stack.
'                       End If
                    Else
                        If JustifiedMiniRowCount(i, 0) > JustifiedMiniRowCount(i, 1) Then                 ' If minirow1 count is greater than minirow2 count, switch minirow1 and minirow2 counts.
                            hold = JustifiedMiniRowCount(i, 0) : JustifiedMiniRowCount(i, 0) = JustifiedMiniRowCount(i, 1) : JustifiedMiniRowCount(i, 1) = hold
                            MiniRowOrderTracker(i, 0) = 1 : MiniRowOrderTracker(i, 1) = 0
                        End If
                        If JustifiedMiniRowCount(i, 1) > JustifiedMiniRowCount(i, 2) Then                 ' If minirow2 count is greater than MiniRow3 count, switch minirow2 and MiniRow3 counts.
                            hold = JustifiedMiniRowCount(i, 1) : JustifiedMiniRowCount(i, 1) = JustifiedMiniRowCount(i, 2) : JustifiedMiniRowCount(i, 2) = hold
                            hold = MiniRowOrderTracker(i, 1) : MiniRowOrderTracker(i, 1) = MiniRowOrderTracker(i, 2) : MiniRowOrderTracker(i, 2) = hold
                        End If
                        If JustifiedMiniRowCount(i, 0) > JustifiedMiniRowCount(i, 1) Then                 ' If minirow1 count is greater than minirow2 count, switch minirow1 and minirow2 counts.
                            hold = JustifiedMiniRowCount(i, 0) : JustifiedMiniRowCount(i, 0) = JustifiedMiniRowCount(i, 1) : JustifiedMiniRowCount(i, 1) = hold
                            hold = MiniRowOrderTracker(i, 0) : MiniRowOrderTracker(i, 0) = MiniRowOrderTracker(i, 1) : MiniRowOrderTracker(i, 1) = hold
                        End If
                        CalcMiniRowCountCode(i) = 16 * JustifiedMiniRowCount(i, 0) + 4 * JustifiedMiniRowCount(i, 1) + JustifiedMiniRowCount(i, 2)
                    End If
                    If CalcMiniRowCountCodeMinimum > CalcMiniRowCountCode(i) Then
                        CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode(i)
                        zstart = i
                    End If
                Next i

                CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = False
                RowWithCalcMiniRowCountCodeMinimumIx = -1
                For i = zstart To 17                                                ' Identify and save the rows that match the minimum count, these are candidates for row1.
                    '                                                                 (Note: All candidates will yield the same row1 after relabeling.)
                    If CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode(i) Then
                        RowWithCalcMiniRowCountCodeMinimumIx += 1
                        RowWithCalcMiniRowCountCodeMinimum(RowWithCalcMiniRowCountCodeMinimumIx) = i    ' NOTE: rows in 0 to 17 notation.
                        If Not CandidatesRepeatDigitsBetweenFirstAndSecondRowSw And CalcMiniRowCountCodeMinimum > 0 Then
                            RowRepeatDigits = DigitsInRowBit(i) And DigitsInRowBit(Row2CandidateA(i))
                            If RowRepeatDigits > 0 Then
                                CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = True
                            Else
                                RowRepeatDigits = DigitsInRowBit(i) And DigitsInRowBit(Row2CandidateB(i))
                                If RowRepeatDigits > 0 Then CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = True
                            End If
                        End If
                    End If
                Next i

                FirstNonZeroPositionInRow1 = FirstNonZeroPositionInFirstNonZeroRow(CalcMiniRowCountCodeMinimum)
                ''Array.ConstrainedCopy(FirstNonZeroRowInit, CalcMiniRowCountCodeMinimum * 9, MinLexGridLocal, 0, 9)   ' Set first row of MinLexGridLocal based on CalcMiniRowCountCodeMinimum.
                'zstart = CalcMiniRowCountCodeMinimum * 9
                'For z = 0 To 8 : MinLexGridLocal(z) = FirstNonZeroRowInit(z + zstart) : Next z
                TwoRowCandidateMiniRowCodeMinimum = 999
                TwoRowCandidateminirow1CodeMinimum = 999
                TwoRowCandidateIx = -1
                Band1CandidateIx = -1
                If CalcMiniRowCountCodeMinimum > 0 Then
                    If CalcMiniRowCountCodeMinimum < 4 Then
                        For i = 0 To RowWithCalcMiniRowCountCodeMinimumIx
                            TwoRowCandidateIx += 1
                            j = RowWithCalcMiniRowCountCodeMinimum(i)
                            k = Row2CandidateA(j)
                            m = MiniRowOrderTracker(j, 0)
                            n = MiniRowOrderTracker(j, 1)
                            If CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, m)
                                If MiniRowRepeatDigits > 0 Then
                                    AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m) * 2 - 1
                                Else
                                    AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m) * 2
                                End If
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, n)
                                If MiniRowRepeatDigits > 0 Then
                                    AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n) * 2 - 1
                                Else
                                    AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n) * 2
                                End If
                            Else
                                AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m)
                                AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n)
                            End If
                            If AugmentedMiniRowCount1 > AugmentedMiniRowCount2 Then
                                TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount2
                                TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount1
                            Else
                                TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount1
                                TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount2
                            End If
                            TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = TwoRowCandidateMiniRow1Count * 10 + TwoRowCandidateMiniRow2Count
                            If TwoRowCandidateMiniRowCodeMinimum > TwoRowCandidateMiniRowCode(TwoRowCandidateIx) Then
                                TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode(TwoRowCandidateIx)
                                zstart = i
                            End If
                            TwoRowCandidateRow1(TwoRowCandidateIx) = j
                            TwoRowCandidateRow2(TwoRowCandidateIx) = k
                            TwoRowCandidateRow3(TwoRowCandidateIx) = Row2CandidateB(j)
                            TwoRowCandidateIx += 1
                            k = Row2CandidateB(j)
                            If CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, m)
                                If MiniRowRepeatDigits > 0 Then
                                    AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m) * 2 - 1
                                Else
                                    AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m) * 2
                                End If
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, n)
                                If MiniRowRepeatDigits > 0 Then
                                    AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n) * 2 - 1
                                Else
                                    AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n) * 2
                                End If
                            Else
                                AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m)
                                AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n)
                            End If
                            If AugmentedMiniRowCount1 > AugmentedMiniRowCount2 Then
                                TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount2
                                TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount1
                            Else
                                TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount1
                                TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount2
                            End If
                            TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = TwoRowCandidateMiniRow1Count * 10 + TwoRowCandidateMiniRow2Count
                            If TwoRowCandidateMiniRowCodeMinimum > TwoRowCandidateMiniRowCode(TwoRowCandidateIx) Then
                                TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode(TwoRowCandidateIx)
                                zstart = i
                            End If
                            TwoRowCandidateRow1(TwoRowCandidateIx) = j
                            TwoRowCandidateRow2(TwoRowCandidateIx) = k
                            TwoRowCandidateRow3(TwoRowCandidateIx) = Row2CandidateA(j)
                        Next i
                        For i = zstart To TwoRowCandidateIx
                            If TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode(i) Then
                                Band1CandidateIx += 1
                                Band1CandidateRow1(Band1CandidateIx) = TwoRowCandidateRow1(i)
                                Band1CandidateRow2(Band1CandidateIx) = TwoRowCandidateRow2(i)
                                Band1CandidateRow3(Band1CandidateIx) = TwoRowCandidateRow3(i)
                                If Band1CandidateRow1(Band1CandidateIx) > 8 Then TranspositionNeededSw = True
                            End If
                        Next i
                    ElseIf CalcMiniRowCountCodeMinimum < 16 Then
                        For i = 0 To RowWithCalcMiniRowCountCodeMinimumIx
                            TwoRowCandidateIx += 1
                            j = RowWithCalcMiniRowCountCodeMinimum(i)
                            k = Row2CandidateA(j)
                            m = MiniRowOrderTracker(j, 0)
                            If CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, m)
                                If MiniRowRepeatDigits > 0 Then
                                    TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = OriginalMiniRowCount(k, m) * 2 - 1
                                Else
                                    TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = OriginalMiniRowCount(k, m) * 2
                                End If
                            Else
                                TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = OriginalMiniRowCount(k, m)
                            End If
                            If TwoRowCandidateminirow1CodeMinimum > TwoRowCandidateMiniRowCode(TwoRowCandidateIx) Then
                                TwoRowCandidateminirow1CodeMinimum = TwoRowCandidateMiniRowCode(TwoRowCandidateIx)
                                zstart = TwoRowCandidateIx
                            End If
                            TwoRowCandidateRow1(TwoRowCandidateIx) = j
                            TwoRowCandidateRow2(TwoRowCandidateIx) = k
                            TwoRowCandidateRow3(TwoRowCandidateIx) = Row2CandidateB(j)
                            TwoRowCandidateIx += 1
                            k = Row2CandidateB(j)
                            If CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, m)
                                If MiniRowRepeatDigits > 0 Then
                                    TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = OriginalMiniRowCount(k, m) * 2 - 1
                                Else
                                    TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = OriginalMiniRowCount(k, m) * 2
                                End If
                            Else
                                TwoRowCandidateMiniRowCode(TwoRowCandidateIx) = OriginalMiniRowCount(k, m)
                            End If
                            If TwoRowCandidateminirow1CodeMinimum > TwoRowCandidateMiniRowCode(TwoRowCandidateIx) Then
                                TwoRowCandidateminirow1CodeMinimum = TwoRowCandidateMiniRowCode(TwoRowCandidateIx)
                                zstart = TwoRowCandidateIx
                            End If
                            TwoRowCandidateRow1(TwoRowCandidateIx) = j
                            TwoRowCandidateRow2(TwoRowCandidateIx) = k
                            TwoRowCandidateRow3(TwoRowCandidateIx) = Row2CandidateA(j)
                        Next i
                        For i = zstart To TwoRowCandidateIx
                            If TwoRowCandidateminirow1CodeMinimum = TwoRowCandidateMiniRowCode(i) Then
                                Band1CandidateIx += 1
                                Band1CandidateRow1(Band1CandidateIx) = TwoRowCandidateRow1(i)
                                Band1CandidateRow2(Band1CandidateIx) = TwoRowCandidateRow2(i)
                                Band1CandidateRow3(Band1CandidateIx) = TwoRowCandidateRow3(i)
                                If Band1CandidateRow1(Band1CandidateIx) > 8 Then TranspositionNeededSw = True
                            End If
                        Next i
                    Else  ' CalcMiniRowCountCodeMinimum > 15
                        For i = 0 To RowWithCalcMiniRowCountCodeMinimumIx
                            z = RowWithCalcMiniRowCountCodeMinimum(i)
                            a = Row2CandidateA(z)
                            b = Row2CandidateB(z)
                            Band1CandidateIx += 1
                            Band1CandidateRow1(Band1CandidateIx) = z
                            Band1CandidateRow2(Band1CandidateIx) = a
                            Band1CandidateRow3(Band1CandidateIx) = b
                            If z > 8 Then TranspositionNeededSw = True
                            Band1CandidateIx += 1
                            Band1CandidateRow1(Band1CandidateIx) = z
                            Band1CandidateRow2(Band1CandidateIx) = b
                            Band1CandidateRow3(Band1CandidateIx) = a
                        Next i
                    End If
                Else ' If rows with all zeros exist, possibly reduce band 1 candidates by evaluating their row2 candidates.
                    ' Identify and save the rows that match the row2 minimum count for all-zeros row1 candidates, these are candidates for row2.
                    ' (Note: All row2 candidates will yield the same row2 after relabeling.)
                    CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = False
                    Row2CalcMiniRowCountCodeMinimum = 999
                    TwoRowCandidateIx = -1
                    For i = 0 To RowWithCalcMiniRowCountCodeMinimumIx
                        z = RowWithCalcMiniRowCountCodeMinimum(i)
                        a = Row2CandidateA(z)
                        b = Row2CandidateB(z)
                        If Row2CalcMiniRowCountCodeMinimum > CalcMiniRowCountCode(a) Then Row2CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode(a) : zstart = i
                        If Row2CalcMiniRowCountCodeMinimum > CalcMiniRowCountCode(b) Then Row2CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode(b) : zstart = i
                    Next i
                    For i = zstart To RowWithCalcMiniRowCountCodeMinimumIx
                        z = RowWithCalcMiniRowCountCodeMinimum(i)
                        a = Row2CandidateA(z)
                        b = Row2CandidateB(z)
                        If Row2CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode(a) Then
                            TwoRowCandidateIx += 1
                            TwoRowCandidateRow1(TwoRowCandidateIx) = z
                            TwoRowCandidateRow2(TwoRowCandidateIx) = a
                            TwoRowCandidateRow3(TwoRowCandidateIx) = b
                            If Not CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                RowRepeatDigits = DigitsInRowBit(a) And DigitsInRowBit(b)
                                If RowRepeatDigits > 0 Then CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = True
                            End If
                        End If
                        If Row2CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode(b) Then
                            TwoRowCandidateIx += 1
                            TwoRowCandidateRow1(TwoRowCandidateIx) = z
                            TwoRowCandidateRow2(TwoRowCandidateIx) = b
                            TwoRowCandidateRow3(TwoRowCandidateIx) = a
                            If Not CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                RowRepeatDigits = DigitsInRowBit(b) And DigitsInRowBit(a)
                                If RowRepeatDigits > 0 Then CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = True
                            End If
                        End If
                    Next i
                    If TwoRowCandidateIx = 0 Then
                        Band1CandidateIx = 0
                        Band1CandidateRow1(0) = TwoRowCandidateRow1(0)
                        Band1CandidateRow2(0) = TwoRowCandidateRow2(0)
                        Band1CandidateRow3(0) = TwoRowCandidateRow3(0)
                        If Band1CandidateRow1(Band1CandidateIx) > 8 Then TranspositionNeededSw = True
                    ElseIf Row2CalcMiniRowCountCodeMinimum < 4 Then
                        For i = 0 To TwoRowCandidateIx
                            j = TwoRowCandidateRow2(i)
                            k = TwoRowCandidateRow3(i)
                            m = MiniRowOrderTracker(j, 0)
                            n = MiniRowOrderTracker(j, 1)
                            If CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, m)
                                If MiniRowRepeatDigits > 0 Then
                                    AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m) * 2 - 1
                                Else
                                    AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m) * 2
                                End If
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, n)
                                If MiniRowRepeatDigits > 0 Then
                                    AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n) * 2 - 1
                                Else
                                    AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n) * 2
                                End If
                            Else
                                AugmentedMiniRowCount1 = OriginalMiniRowCount(k, m)
                                AugmentedMiniRowCount2 = OriginalMiniRowCount(k, n)
                            End If
                            If AugmentedMiniRowCount1 > AugmentedMiniRowCount2 Then
                                TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount2
                                TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount1
                            Else
                                TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount1
                                TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount2
                            End If
                            TwoRowCandidateMiniRowCode(i) = TwoRowCandidateMiniRow1Count * 10 + TwoRowCandidateMiniRow2Count
                            If TwoRowCandidateMiniRowCodeMinimum > TwoRowCandidateMiniRowCode(i) Then
                                TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode(i)
                                zstart = i
                            End If
                        Next i
                        For i = zstart To TwoRowCandidateIx
                            If TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode(i) Then
                                Band1CandidateIx += 1
                                Band1CandidateRow1(Band1CandidateIx) = TwoRowCandidateRow1(i)
                                Band1CandidateRow2(Band1CandidateIx) = TwoRowCandidateRow2(i)
                                Band1CandidateRow3(Band1CandidateIx) = TwoRowCandidateRow3(i)
                                If Band1CandidateRow1(Band1CandidateIx) > 8 Then TranspositionNeededSw = True
                            End If
                        Next i
                    ElseIf Row2CalcMiniRowCountCodeMinimum < 16 Then
                        For i = 0 To TwoRowCandidateIx
                            j = TwoRowCandidateRow2(i)
                            k = TwoRowCandidateRow3(i)
                            m = MiniRowOrderTracker(j, 0)
                            If CandidatesRepeatDigitsBetweenFirstAndSecondRowSw Then
                                MiniRowRepeatDigits = DigitsInRowBit(j) And DigitsInMiniRowBit(k, m)
                                If MiniRowRepeatDigits > 0 Then
                                    TwoRowCandidateMiniRowCode(i) = OriginalMiniRowCount(k, m) * 2 - 1
                                Else
                                    TwoRowCandidateMiniRowCode(i) = OriginalMiniRowCount(k, m) * 2
                                End If
                            Else
                                TwoRowCandidateMiniRowCode(i) = OriginalMiniRowCount(k, m)
                            End If
                            If TwoRowCandidateminirow1CodeMinimum > TwoRowCandidateMiniRowCode(i) Then
                                TwoRowCandidateminirow1CodeMinimum = TwoRowCandidateMiniRowCode(i)
                                zstart = i
                            End If
                        Next i
                        For i = zstart To TwoRowCandidateIx
                            If TwoRowCandidateminirow1CodeMinimum = TwoRowCandidateMiniRowCode(i) Then
                                Band1CandidateIx += 1
                                Band1CandidateRow1(Band1CandidateIx) = TwoRowCandidateRow1(i)
                                Band1CandidateRow2(Band1CandidateIx) = TwoRowCandidateRow2(i)
                                Band1CandidateRow3(Band1CandidateIx) = TwoRowCandidateRow3(i)
                                If Band1CandidateRow1(Band1CandidateIx) > 8 Then TranspositionNeededSw = True
                            End If
                        Next i
                    Else ' Row2CalcMiniRowCountCodeMinimum > 15
                        For i = 0 To TwoRowCandidateIx
                            Band1CandidateIx += 1
                            Band1CandidateRow1(Band1CandidateIx) = TwoRowCandidateRow1(i)
                            Band1CandidateRow2(Band1CandidateIx) = TwoRowCandidateRow2(i)
                            Band1CandidateRow3(Band1CandidateIx) = TwoRowCandidateRow3(i)
                            If TwoRowCandidateRow1(i) > 8 Then TranspositionNeededSw = True
                        Next i
                    End If
                    FirstNonZeroPositionInRow2 = FirstNonZeroPositionInFirstNonZeroRow(Row2CalcMiniRowCountCodeMinimum)
                    ''Array.ConstrainedCopy(FirstNonZeroRowInit, Row2CalcMiniRowCountCodeMinimum * 9, MinLexGridLocal, 9, 9)   ' For cases of all zeros row1, Set the second row of MinLexGridLocal based on Row2CalcMiniRowCountCodeMinimum.
                    'zstart = Row2CalcMiniRowCountCodeMinimum * 9
                    'For z = 0 To 8 : MinLexGridLocal(z + 9) = FirstNonZeroRowInit(z + zstart) : Next z
                End If

                If TranspositionNeededSw Then
                    For i = 0 To 80                                               ' transpose rows to columns.
                        InputGrid(81 + i \ 9 + (i Mod 9) * 9) = InputGrid(i)
                    Next i
                End If

                ' Band1 Processing - Two cases:
                '    1) MinLexBand1:     If all 18 rows and columns have at least one non-zero digit, then MinLex Band1 for all row1 candidates and eliminate Band1 row/column, permutation & relabel possibilities that do not yield the Band1 MinLex.
                '    2) MinLexRows2and3: If all-zeros rows or columns exist, then MinLex Band1 for all row2 candidates for the all-zeros row and columns and eliminate Band1 row/column, permutation & relabel possibilities that do not yield the Band1 MinLex.
                '    Results:            1) MinLexed Band1 solution candidates ready for rows 4 - 9 processing.
                '                        2) Column (and stack) permutation Trackers used to reproduce each puzzle configuration that yields its Band1 MinLex.
                '                        3) Relabel Trackers for each Band1 MinLex configuration.

                StoppedJustifyingRow = 0
                If CalcMiniRowCountCodeMinimum > 0 Then      ' Start MinLexBand1 (Row1 not empty case)
                    FirstNonZeroDigitPositionInRow(2) = -1
                    Row1StackPermutationCode = FirstNonZeroRowStackPermutationCode(CalcMiniRowCountCodeMinimum)
                    Array.ConstrainedCopy(MinLexGridLocal, 0, MinLexRows, 0, 27)     ' Initialize MinLexRows (Band1 MinLex work array) to all 10s.
                    ColumnPermutationTrackerIx = -1
                    For candidate = 0 To Band1CandidateIx
                        Row1Candidate17 = Band1CandidateRow1(candidate)
                        If Row1Candidate17 < 9 Then
                            Row2Candidate17 = Band1CandidateRow2(candidate)
                            Row1Candidate = Row1Candidate17 + 1
                            Row2Candidate = Band1CandidateRow2(candidate) + 1
                            Row3Candidate = Band1CandidateRow3(candidate) + 1
                            Array.ConstrainedCopy(InputGrid, 0, HoldPuzzle, 0, 81)   ' Copy the direct input puzzle to HoldPuzzle.
                        Else
                            Row2Candidate17 = Band1CandidateRow2(candidate)
                            Row1Candidate = Row1Candidate17 - 8
                            Row2Candidate = Band1CandidateRow2(candidate) - 8
                            Row3Candidate = Band1CandidateRow3(candidate) - 8
                            Array.ConstrainedCopy(InputGrid, 81, HoldPuzzle, 0, 81)  ' Copy the transposed input puzzle to HoldPuzzle.
                        End If
                        ' MinLex Band1 for Row1Candidate.
                        MinLexRowHitSw = False
                        ResetMinLexRowSw = False
                        CandidateColumnPermutationTrackerStartIx = ColumnPermutationTrackerIx
                        row1start = (Row1Candidate - 1) * 9
                        row2start = (Row2Candidate - 1) * 9
                        row3start = (Row3Candidate - 1) * 9

                        StackPermutationTrackerLocal(0) = MiniRowOrderTracker(Row1Candidate17, 0)
                        StackPermutationTrackerLocal(1) = MiniRowOrderTracker(Row1Candidate17, 1)
                        StackPermutationTrackerLocal(2) = MiniRowOrderTracker(Row1Candidate17, 2)
                        Row1MiniRowCount(0) = JustifiedMiniRowCount(Row1Candidate17, 0)
                        Row1MiniRowCount(1) = JustifiedMiniRowCount(Row1Candidate17, 1)
                        Row1MiniRowCount(2) = JustifiedMiniRowCount(Row1Candidate17, 2)
                        Row2MiniRowCount(0) = OriginalMiniRowCount(Row2Candidate17, StackPermutationTrackerLocal(0))
                        Row2MiniRowCount(1) = OriginalMiniRowCount(Row2Candidate17, StackPermutationTrackerLocal(1))
                        Row2MiniRowCount(2) = OriginalMiniRowCount(Row2Candidate17, StackPermutationTrackerLocal(2))

                        Row2Or3StackPermutationCode = Row1StackPermutationCode
                        If CalcMiniRowCountCodeMinimum < 4 Or Row1StackPermutationCode > 1 Then
                            If Row2MiniRowCount(0) > Row2MiniRowCount(1) Then      ' If Row2 minirow1 count is greater than minirow2, switch minirow1 and minirow2.
                                hold = Row2MiniRowCount(0) : Row2MiniRowCount(0) = Row2MiniRowCount(1) : Row2MiniRowCount(1) = hold
                                hold = StackPermutationTrackerLocal(0) : StackPermutationTrackerLocal(0) = StackPermutationTrackerLocal(1) : StackPermutationTrackerLocal(1) = hold
                            ElseIf CalcMiniRowCountCodeMinimum < 4 AndAlso Row2MiniRowCount(1) > 0 And Row2MiniRowCount(0) = Row2MiniRowCount(1) Then
                                Row2Or3StackPermutationCode = 2
                            End If
                        End If

                        MinLexCandidateSw = True
                        If Row1StackPermutationCode = 0 Then
                            Select Case FirstNonZeroDigitPositionInRow(2)
                                Case < 1    ' If -1 or 0, do nothing.
                                Case 5
                                    If Row2MiniRowCount(0) > 0 Or Row2MiniRowCount(1) > 1 Then MinLexCandidateSw = False
                                Case 4
                                    If Row2MiniRowCount(0) > 0 Or Row2MiniRowCount(1) = 3 Then MinLexCandidateSw = False
                                Case > 5
                                    If Row2MiniRowCount(0) > 0 Or Row2MiniRowCount(1) > 0 Then MinLexCandidateSw = False
                                Case 3
                                    If Row2MiniRowCount(0) > 0 Then MinLexCandidateSw = False
                                Case 2
                                    If Row2MiniRowCount(0) > 1 Then MinLexCandidateSw = False
                                Case 1
                                    If Row2MiniRowCount(0) = 3 Then MinLexCandidateSw = False
                            End Select
                        End If

                        If MinLexCandidateSw Then
                            StackPermutationTrackerCode = StackPermutationTrackerLocal(0) * 100 + StackPermutationTrackerLocal(1) * 10 + StackPermutationTrackerLocal(2)
                            Select Case StackPermutationTrackerCode        ' Apply permutations required to right justify row 1 and row 2.
                                Case 12     ' 012    - no permutation
                                Case 21     ' 021
                                    SwitchStacks12(HoldPuzzle)
                                Case 102
                                    SwitchStacks01(HoldPuzzle)
                                Case 120
                                    Switch3Stacks120(HoldPuzzle)
                                Case 201
                                    Switch3Stacks201(HoldPuzzle)
                                Case 210
                                    SwitchStacks02(HoldPuzzle)
                            End Select

                            Row3MiniRowCount(0) = 0
                            Row3MiniRowCount(1) = 0
                            Row3MiniRowCount(2) = 0
                            For i = 0 To 8                                             ' Count row 3 non-zero digits in MiniRows 1, 2 and 3.
                                If HoldPuzzle(row3start + i) > 0 Then
                                    Row3MiniRowCount(i \ 3) += 1
                                End If
                            Next i
                            If Row2MiniRowCount(0) = 0 And Row2MiniRowCount(1) = 0 Then
                                If Row3MiniRowCount(0) > Row3MiniRowCount(1) Then      ' If row 3 minirow1 count is greater than minirow2, switch minirow1 and minirow2.
                                    hold = Row3MiniRowCount(0) : Row3MiniRowCount(0) = Row3MiniRowCount(1) : Row3MiniRowCount(1) = hold
                                    SwitchStacks01(HoldPuzzle)
                                ElseIf Row3MiniRowCount(1) > 0 And Row3MiniRowCount(0) = Row3MiniRowCount(1) Then
                                    Row2Or3StackPermutationCode = 2
                                End If
                            End If

                            ' Positionally (not considering digit values) "right justify" first, second and third row non-zero digits within MiniRows.
                            ' Also set the ColumnPermutationCode for Band1 using MiniRow 0 to 2 notation.
                            Band1MiniRowColumnPermutationCode(0) = 0
                            Band1MiniRowColumnPermutationCode(1) = 0
                            Band1MiniRowColumnPermutationCode(2) = 0
                            Select Case Row1MiniRowCount(0)                               ' MiniRow 1
                                Case 0   ' Row1MiniRowCount(0) = 0
                                    Select Case Row2MiniRowCount(0)
                                        Case 0 ' Row 2
                                            Select Case Row3MiniRowCount(0)
                                                Case 0   ' Do nothing.
                                                Case 1
                                                    If HoldPuzzle(row3start) > 0 Then
                                                        SwitchColumns02(HoldPuzzle)
                                                    ElseIf HoldPuzzle(row3start + 1) > 0 Then
                                                        SwitchColumns12(HoldPuzzle)
                                                    End If
                                                Case 2
                                                    If HoldPuzzle(row3start + 2) = 0 Then
                                                        SwitchColumns02(HoldPuzzle)
                                                    ElseIf HoldPuzzle(row3start + 1) = 0 Then
                                                        SwitchColumns01(HoldPuzzle)
                                                    End If
                                                    Band1MiniRowColumnPermutationCode(0) += 1
                                                Case 3
                                                    Band1MiniRowColumnPermutationCode(0) += 3
                                            End Select
                                        Case 1 ' Row2
                                            If HoldPuzzle(row2start) > 0 Then
                                                SwitchColumns02(HoldPuzzle)
                                            ElseIf HoldPuzzle(row2start + 1) > 0 Then
                                                SwitchColumns12(HoldPuzzle)
                                            End If
                                            If HoldPuzzle(row3start) > 0 Then
                                                If HoldPuzzle(row3start + 1) = 0 Then
                                                    SwitchColumns01(HoldPuzzle)
                                                Else
                                                    Band1MiniRowColumnPermutationCode(0) += 2
                                                End If
                                            End If
                                        Case 2 ' Row2
                                            If HoldPuzzle(row2start + 2) = 0 Then
                                                SwitchColumns02(HoldPuzzle)
                                            ElseIf HoldPuzzle(row2start + 1) = 0 Then
                                                SwitchColumns01(HoldPuzzle)
                                            End If
                                            Band1MiniRowColumnPermutationCode(0) += 1
                                        Case 3 ' Row2
                                            Band1MiniRowColumnPermutationCode(0) += 3
                                    End Select
                                Case 1 ' Row1MiniRowCount(0) = 1
                                    If HoldPuzzle(row1start) > 0 Then
                                        SwitchColumns02(HoldPuzzle)
                                    ElseIf HoldPuzzle(row1start + 1) > 0 Then
                                        SwitchColumns12(HoldPuzzle)
                                    End If
                                    If HoldPuzzle(row2start) > 0 Then
                                        If HoldPuzzle(row2start + 1) = 0 Then
                                            SwitchColumns01(HoldPuzzle)
                                        Else
                                            Band1MiniRowColumnPermutationCode(0) += 2
                                        End If
                                    Else
                                        If HoldPuzzle(row2start + 1) = 0 Then
                                            If HoldPuzzle(row3start) > 0 Then
                                                If HoldPuzzle(row3start + 1) = 0 Then
                                                    SwitchColumns01(HoldPuzzle)
                                                Else
                                                    Band1MiniRowColumnPermutationCode(0) += 2
                                                End If
                                            End If
                                        End If
                                    End If
                                Case 2    ' Row1MiniRowCount(0) = 2
                                    If HoldPuzzle(row1start + 2) = 0 Then
                                        SwitchColumns02(HoldPuzzle)
                                    ElseIf HoldPuzzle(row1start + 1) = 0 Then
                                        SwitchColumns01(HoldPuzzle)
                                    End If
                                    Band1MiniRowColumnPermutationCode(0) += 1
                                Case 3    ' Row1MiniRowCount(0) = 3
                                    Band1MiniRowColumnPermutationCode(0) += 3
                            End Select

                            Select Case Row1MiniRowCount(1)                               ' MiniRow 2
                                Case 0   ' Row1MiniRowCount(1) = 0
                                    Select Case Row2MiniRowCount(1)
                                        Case 0 ' Row 2
                                            Select Case Row3MiniRowCount(1)
                                                Case 0   ' Do nothing.
                                                Case 1
                                                    If HoldPuzzle(row3start + 3) > 0 Then
                                                        SwitchColumns35(HoldPuzzle)
                                                    ElseIf HoldPuzzle(row3start + 4) > 0 Then
                                                        SwitchColumns45(HoldPuzzle)
                                                    End If
                                                Case 2
                                                    If HoldPuzzle(row3start + 5) = 0 Then
                                                        SwitchColumns35(HoldPuzzle)
                                                    ElseIf HoldPuzzle(row3start + 4) = 0 Then
                                                        SwitchColumns34(HoldPuzzle)
                                                    End If
                                                    Band1MiniRowColumnPermutationCode(1) += 1
                                                Case 3
                                                    Band1MiniRowColumnPermutationCode(1) += 3
                                            End Select
                                        Case 1 ' Row2
                                            If HoldPuzzle(row2start + 3) > 0 Then
                                                SwitchColumns35(HoldPuzzle)
                                            ElseIf HoldPuzzle(row2start + 4) > 0 Then
                                                SwitchColumns45(HoldPuzzle)
                                            End If
                                            If HoldPuzzle(row3start + 3) > 0 Then
                                                If HoldPuzzle(row3start + 4) = 0 Then
                                                    SwitchColumns34(HoldPuzzle)
                                                Else
                                                    Band1MiniRowColumnPermutationCode(1) += 2
                                                End If
                                            End If
                                        Case 2 ' Row2
                                            If HoldPuzzle(row2start + 5) = 0 Then
                                                SwitchColumns35(HoldPuzzle)
                                            ElseIf HoldPuzzle(row2start + 4) = 0 Then
                                                SwitchColumns34(HoldPuzzle)
                                            End If
                                            Band1MiniRowColumnPermutationCode(1) += 1
                                        Case 3 ' Row2
                                            Band1MiniRowColumnPermutationCode(1) += 3
                                    End Select
                                Case 1   ' Row1MiniRowCount(1) = 1
                                    If HoldPuzzle(row1start + 3) > 0 Then
                                        SwitchColumns35(HoldPuzzle)
                                    ElseIf HoldPuzzle(row1start + 4) > 0 Then
                                        SwitchColumns45(HoldPuzzle)
                                    End If
                                    If HoldPuzzle(row2start + 3) > 0 Then
                                        If HoldPuzzle(row2start + 4) = 0 Then
                                            SwitchColumns34(HoldPuzzle)
                                        Else
                                            Band1MiniRowColumnPermutationCode(1) += 2
                                        End If
                                    Else
                                        If HoldPuzzle(row2start + 4) = 0 Then
                                            If HoldPuzzle(row3start + 3) > 0 Then
                                                If HoldPuzzle(row3start + 4) = 0 Then
                                                    SwitchColumns34(HoldPuzzle)
                                                Else
                                                    Band1MiniRowColumnPermutationCode(1) += 2
                                                End If
                                            End If
                                        End If
                                    End If
                                Case 2   ' Row1MiniRowCount(1) = 2
                                    If HoldPuzzle(row1start + 5) = 0 Then
                                        SwitchColumns35(HoldPuzzle)
                                    ElseIf HoldPuzzle(row1start + 4) = 0 Then
                                        SwitchColumns34(HoldPuzzle)
                                    End If
                                    Band1MiniRowColumnPermutationCode(1) += 1
                                Case 3   ' Row1MiniRowCount(1) = 3
                                    Band1MiniRowColumnPermutationCode(1) += 3
                            End Select

                            Select Case Row1MiniRowCount(2)                               ' MiniRow 3
                            ' Case 0 not possible since this section handles non-zero row1 candidates.
                                Case 1   ' Row1MiniRowCount(2) = 1
                                    If HoldPuzzle(row1start + 6) > 0 Then
                                        SwitchColumns68(HoldPuzzle)
                                    ElseIf HoldPuzzle(row1start + 7) > 0 Then
                                        SwitchColumns78(HoldPuzzle)
                                    End If
                                    If HoldPuzzle(row2start + 6) > 0 Then
                                        If HoldPuzzle(row2start + 7) = 0 Then
                                            SwitchColumns67(HoldPuzzle)
                                        Else
                                            Band1MiniRowColumnPermutationCode(2) += 2
                                        End If
                                    Else
                                        If HoldPuzzle(row2start + 7) = 0 Then
                                            If HoldPuzzle(row3start + 6) > 0 Then
                                                If HoldPuzzle(row3start + 7) = 0 Then
                                                    SwitchColumns67(HoldPuzzle)
                                                Else
                                                    Band1MiniRowColumnPermutationCode(2) += 2
                                                End If
                                            End If
                                        End If
                                    End If
                                Case 2   ' Row1MiniRowCount(2) = 2
                                    If HoldPuzzle(row1start + 8) = 0 Then
                                        SwitchColumns68(HoldPuzzle)
                                    ElseIf HoldPuzzle(row1start + 7) = 0 Then
                                        SwitchColumns67(HoldPuzzle)
                                    End If
                                    Band1MiniRowColumnPermutationCode(2) += 1
                                Case 3   ' Row1MiniRowCount(2) = 3
                                    Band1MiniRowColumnPermutationCode(2) += 3
                            End Select

                            'Array.ConstrainedCopy(HoldPuzzle, (Row1Candidate - 1) * 9, LocalBand1, 0, 9)
                            zstart = (Row1Candidate - 1) * 9
                            For z = 0 To 8 : LocalBand1(z) = HoldPuzzle(z + zstart) : Next z
                            'Array.ConstrainedCopy(HoldPuzzle, (Row2Candidate - 1) * 9, LocalBand1, 9, 9)
                            zstart = (Row2Candidate - 1) * 9
                            For z = 0 To 8 : LocalBand1(z + 9) = HoldPuzzle(z + zstart) : Next z
                            'Array.ConstrainedCopy(HoldPuzzle, (Row3Candidate - 1) * 9, LocalBand1, 18, 9)
                            zstart = (Row3Candidate - 1) * 9
                            For z = 0 To 8 : LocalBand1(z + 18) = HoldPuzzle(z + zstart) : Next z
                            'Array.ConstrainedCopy(ColumnTrackerInit, 0, LocalColumnPermutationTracker, 0, 9)
                            For z = 0 To 8 : LocalColumnPermutationTracker(z) = ColumnTrackerInit(z) : Next z
                            FirstColumnPermutationTrackerIsIdentitySw = True
                            For stackpermutationix = 0 To StackPermutations(Row2Or3StackPermutationCode)
                                If stackpermutationix > 0 Then
                                    FirstColumnPermutationTrackerIsIdentitySw = False
                                    stackx = PermutationStackX(Row2Or3StackPermutationCode)(stackpermutationix)          ' Band1 (3 row) stack switch.
                                    stacky = PermutationStackY(Row2Or3StackPermutationCode)(stackpermutationix)
                                    switchx = 3 * stackx
                                    switchy = 3 * stacky
                                    hold = Band1MiniRowColumnPermutationCode(stackx)
                                    Band1MiniRowColumnPermutationCode(stackx) = Band1MiniRowColumnPermutationCode(stacky)
                                    Band1MiniRowColumnPermutationCode(stacky) = hold
                                    hold = Row2MiniRowCount(stackx)
                                    Row2MiniRowCount(stackx) = Row2MiniRowCount(stacky)
                                    Row2MiniRowCount(stacky) = hold
                                    hold = LocalColumnPermutationTracker(switchx)
                                    LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                    LocalColumnPermutationTracker(switchy) = hold
                                    hold = LocalBand1(switchx)
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalColumnPermutationTracker(switchx)
                                    LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                    LocalColumnPermutationTracker(switchy) = hold
                                    hold = LocalBand1(switchx)
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalColumnPermutationTracker(switchx)
                                    LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                    LocalColumnPermutationTracker(switchy) = hold
                                    hold = LocalBand1(switchx)
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 7
                                    switchy += 7
                                    hold = LocalBand1(switchx)                                                     ' row 2
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalBand1(switchx)
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalBand1(switchx)
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 7
                                    switchy += 7
                                    hold = LocalBand1(switchx)                                                     ' row 3
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalBand1(switchx)
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalBand1(switchx)
                                    LocalBand1(switchx) = LocalBand1(switchy)
                                    LocalBand1(switchy) = hold
                                End If

                                MinLexCandidateSw = True
                                Select Case FirstNonZeroDigitPositionInRow(2)
                                    Case < 1    ' If -1 or 0, do nothing.
                                    Case 5
                                        If Row2MiniRowCount(0) > 0 Or LocalBand1(12) > 0 Or LocalBand1(13) > 0 Then MinLexCandidateSw = False
                                    Case 4
                                        If Row2MiniRowCount(0) > 0 Or LocalBand1(12) > 0 Then MinLexCandidateSw = False
                                    Case > 5
                                        If Row2MiniRowCount(0) > 0 Or Row2MiniRowCount(1) > 0 Then MinLexCandidateSw = False
                                    Case 3
                                        If Row2MiniRowCount(0) > 0 Then MinLexCandidateSw = False
                                    Case 2
                                        If LocalBand1(9) > 0 Or LocalBand1(10) > 0 Then MinLexCandidateSw = False
                                    Case 1
                                        If LocalBand1(9) > 0 Then MinLexCandidateSw = False
                                End Select

                                If MinLexCandidateSw Then
                                    ColumnPermutationCode = Band1MiniRowColumnPermutationCode(0) * 16 + Band1MiniRowColumnPermutationCode(1) * 4 + Band1MiniRowColumnPermutationCode(2)
                                    For columnpermutationix = 0 To ColumnPermutations(ColumnPermutationCode)
                                        If columnpermutationix > 0 Then                    ' Band1 (3 row) column permutation.
                                            FirstColumnPermutationTrackerIsIdentitySw = False
                                            switchx = PermutationColumnX(ColumnPermutationCode)(columnpermutationix)
                                            switchy = PermutationColumnY(ColumnPermutationCode)(columnpermutationix)
                                            hold = LocalColumnPermutationTracker(switchx)
                                            LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                            LocalColumnPermutationTracker(switchy) = hold
                                            hold = LocalBand1(switchx)                     ' row 1
                                            LocalBand1(switchx) = LocalBand1(switchy)
                                            LocalBand1(switchy) = hold
                                            switchx += 9                                   ' row2
                                            switchy += 9
                                            hold = LocalBand1(switchx)
                                            LocalBand1(switchx) = LocalBand1(switchy)
                                            LocalBand1(switchy) = hold
                                            switchx += 9                                   ' row3
                                            switchy += 9
                                            hold = LocalBand1(switchx)
                                            LocalBand1(switchx) = LocalBand1(switchy)
                                            LocalBand1(switchy) = hold
                                        Else
                                            For i = 9 To 17
                                                If LocalBand1(i) > 0 Then
                                                    Row2TestFirstNonZeroDigitPositionInRow = i - 9                      ' Note: 0-8 notation.
                                                    If FirstNonZeroDigitPositionInRow(2) < Row2TestFirstNonZeroDigitPositionInRow Then
                                                        FirstNonZeroDigitPositionInRow(2) = Row2TestFirstNonZeroDigitPositionInRow
                                                    End If
                                                    Exit For
                                                End If
                                            Next i
                                        End If

                                        If FirstNonZeroDigitPositionInRow(2) = Row2TestFirstNonZeroDigitPositionInRow Then   ' Check if can bypass relable check is no repeating digit in two rows. ??????
                                            'Array.Clear(DigitAlreadyHitSw, 0, 10)
                                            For z = 0 To 9 : DigitAlreadyHitSw(z) = False : Next z
                                            'Array.Clear(TestRowsRelabeled, 0,  FirstNonZeroPositionInRow + 1)
                                            For z = 0 To FirstNonZeroPositionInRow1 : TestRowsRelabeled(z) = 0 : Next z
                                            'Array.ConstrainedCopy(DigitsRelabelWrkInit, 0, DigitsRelabelWrk, 0, 10)      ' Initialize DigitsRelabelWrk 1 to 9 to 10's (the zero element = 0).
                                            For z = 0 To 9 : DigitsRelabelWrk(z) = DigitsRelabelWrkInit(z) : Next z
                                            RelabelLastDigit = 0
                                            For i = FirstNonZeroPositionInRow1 To 26                                       ' Build DigitsRelabelWrk and TestGridRelabeled for Band1
                                                If LocalBand1(i) > 0 AndAlso Not DigitAlreadyHitSw(LocalBand1(i)) Then
                                                    DigitAlreadyHitSw(LocalBand1(i)) = True
                                                    RelabelLastDigit += 1
                                                    DigitsRelabelWrk(LocalBand1(i)) = RelabelLastDigit
                                                End If
                                                TestRowsRelabeled(i) = DigitsRelabelWrk(LocalBand1(i))
                                            Next i

                                            For i = FirstNonZeroPositionInRow1 To 26
                                                If TestRowsRelabeled(i) > MinLexRows(i) Then   ' Check if row2 is a candidate.
                                                    MinLexCandidateSw = False
                                                    Exit For
                                                ElseIf TestRowsRelabeled(i) < MinLexRows(i) Then
                                                    Exit For
                                                End If
                                            Next i
                                            If MinLexCandidateSw Then
                                                If i < 27 Then
                                                    Array.ConstrainedCopy(TestRowsRelabeled, 0, MinLexRows, 0, 27)
                                                    MinLexRowHitSw = False
                                                    ResetMinLexRowSw = True
                                                    ColumnPermutationTrackerIx = 0
                                                    ColumnPermutationTrackerIsIdentitySw(0) = FirstColumnPermutationTrackerIsIdentitySw
                                                    'Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, 0, 9)
                                                    For z = 0 To 8 : ColumnPermutationTracker(z) = LocalColumnPermutationTracker(z) : Next z
                                                    'Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, 0, 10)
                                                    zstart = ColumnPermutationTrackerIx * 10
                                                    For z = 0 To 9 : HoldDigitsRelabelWrk(z + zstart) = DigitsRelabelWrk(z) : Next z
                                                Else
                                                    MinLexRowHitSw = True
                                                    ColumnPermutationTrackerIx += 1
                                                    If FoundColumnPermutationTrackerCountMax < ColumnPermutationTrackerIx + 1 Then
                                                        FoundColumnPermutationTrackerCountMax = ColumnPermutationTrackerIx + 1
                                                    End If
'                                                   If ColumnPermutationTrackerArrayMax < ColumnPermutationTrackerIx Then
'                                                       ErrorInputRecordStr = Nothing
'                                                       For j = 0 To 80
'                                                           ErrorInputRecordStr += CStr(InputGridBufferChr(inputgridstart + j))
'                                                       Next j
'                                                       Return 3    ' Error 3, Too many column permutations for tracker array.
'                                                   Else
                                                        ColumnPermutationTrackerIsIdentitySw(ColumnPermutationTrackerIx) = FirstColumnPermutationTrackerIsIdentitySw
                                                        'Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, ColumnPermutationTrackerIx * 9, 9)
                                                        zstart = ColumnPermutationTrackerIx * 9
                                                        For z = 0 To 8 : ColumnPermutationTracker(z + zstart) = LocalColumnPermutationTracker(z) : Next z
                                                        'Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, ColumnPermutationTrackerIx * 10, 10)
                                                        zstart = ColumnPermutationTrackerIx * 10
                                                        For z = 0 To 9 : HoldDigitsRelabelWrk(z + zstart) = DigitsRelabelWrk(z) : Next z
'                                                   End If
                                                End If
                                            End If ' If MinLexCandidateSw
                                        End If ' If FirstNonZeroDigitPositionInRow(2) = Row2TestFirstNonZeroDigitPositionInRow
                                        MinLexCandidateSw = True
                                    Next columnpermutationix
                                End If ' If MinLexCandidateSw
                            Next stackpermutationix
                        End If ' If MinLexCandidateSw

                        If ResetMinLexRowSw Then
                            Step2Row1CandidateIx = 0
                            Step2Row1Candidate(0) = Row1Candidate
                            Step2Row2Candidate(0) = Row2Candidate
                            Step2ColumnPermutationTrackerStartIx(0) = -1
                            Step2ColumnPermutationTrackerCount(0) = ColumnPermutationTrackerIx + 1
                            Array.ConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, 0, 81)
                            'Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, 0, 10)
                            For z = 0 To 9 : HoldDigitAlreadyHitSw(z) = DigitAlreadyHitSw(z) : Next z
                            HoldRelabelLastDigit(0) = RelabelLastDigit
                        ElseIf MinLexRowHitSw Then
                            Step2Row1CandidateIx += 1
                            Step2Row1Candidate(Step2Row1CandidateIx) = Row1Candidate
                            Step2Row2Candidate(Step2Row1CandidateIx) = Row2Candidate
                            Step2ColumnPermutationTrackerStartIx(Step2Row1CandidateIx) = CandidateColumnPermutationTrackerStartIx
                            Step2ColumnPermutationTrackerCount(Step2Row1CandidateIx) = ColumnPermutationTrackerIx - CandidateColumnPermutationTrackerStartIx
                            Array.ConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, Step2Row1CandidateIx * 81, 81)
                            'Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, Step2Row1CandidateIx * 10, 10)
                            zstart = Step2Row1CandidateIx * 10
                            For z = 0 To 9 : HoldDigitAlreadyHitSw(z + zstart) = DigitAlreadyHitSw(z) : Next z
                            HoldRelabelLastDigit(Step2Row1CandidateIx) = RelabelLastDigit
                        End If

                    Next candidate
                    Array.ConstrainedCopy(MinLexRows, 0, MinLexGridLocal, 0, 27)   ' For cases of non-zero row1, Copy MinLexRows to MinLexGridLocal.
                    ' End MinLexBand1 (Row1 not empty case)
                Else     ' Start MinLexRows2and3 (Row1 empty case)
                    FirstNonZeroDigitPositionInRow(3) = -1
                    Row2StackPermutationCode = FirstNonZeroRowStackPermutationCode(Row2CalcMiniRowCountCodeMinimum)
                    'Array.ConstrainedCopy(MinLexGridLocal, 9, MinLexRows, 0, 18)                               ' Copy first non-zero row to MinLexRows row 1; set second row to all 10s.
                    For z = 0 To 17 : MinLexRows(z) = MinLexGridLocal(z + 9) : Next z
                    ColumnPermutationTrackerIx = -1
                    For candidate = 0 To Band1CandidateIx
                        Row1Candidate17 = Band1CandidateRow1(candidate)
                        If Row1Candidate17 < 9 Then
                            Row2Candidate17 = Band1CandidateRow2(candidate)
                            Row3Candidate17 = Band1CandidateRow3(candidate)
                            Row1Candidate = Row1Candidate17 + 1
                            Row2Candidate = Band1CandidateRow2(candidate) + 1
                            Row3Candidate = Band1CandidateRow3(candidate) + 1
                            Array.ConstrainedCopy(InputGrid, 0, HoldPuzzle, 0, 81)   ' Copy the direct input puzzle to HoldPuzzle.
                        Else
                            Row2Candidate17 = Band1CandidateRow2(candidate)
                            Row3Candidate17 = Band1CandidateRow3(candidate)
                            Row1Candidate = Row1Candidate17 - 8
                            Row2Candidate = Band1CandidateRow2(candidate) - 8
                            Row3Candidate = Band1CandidateRow3(candidate) - 8
                            Array.ConstrainedCopy(InputGrid, 81, HoldPuzzle, 0, 81)   ' Copy the transposed input puzzle to HoldPuzzle.
                        End If

                        ' MinLex Rows 2 & 3 for Row2Candidate.

                        MinLexRowHitSw = False
                        ResetMinLexRowSw = False
                        CandidateColumnPermutationTrackerStartIx = ColumnPermutationTrackerIx
                        row1start = (Row1Candidate - 1) * 9
                        row2start = (Row2Candidate - 1) * 9
                        row3start = (Row3Candidate - 1) * 9

                        StackPermutationTrackerLocal(0) = MiniRowOrderTracker(Row2Candidate17, 0)
                        StackPermutationTrackerLocal(1) = MiniRowOrderTracker(Row2Candidate17, 1)
                        StackPermutationTrackerLocal(2) = MiniRowOrderTracker(Row2Candidate17, 2)
                        Row2MiniRowCount(0) = JustifiedMiniRowCount(Row2Candidate17, 0)
                        Row2MiniRowCount(1) = JustifiedMiniRowCount(Row2Candidate17, 1)
                        Row2MiniRowCount(2) = JustifiedMiniRowCount(Row2Candidate17, 2)
                        Row3MiniRowCount(0) = OriginalMiniRowCount(Row3Candidate17, StackPermutationTrackerLocal(0))
                        Row3MiniRowCount(1) = OriginalMiniRowCount(Row3Candidate17, StackPermutationTrackerLocal(1))
                        Row3MiniRowCount(2) = OriginalMiniRowCount(Row3Candidate17, StackPermutationTrackerLocal(2))

                        Row3StackPermutationCode = Row2StackPermutationCode
                        If Row2CalcMiniRowCountCodeMinimum < 4 Or Row2StackPermutationCode > 1 Then
                            If Row3MiniRowCount(0) > Row3MiniRowCount(1) Then      ' If Row3 minirow1 count is greater than minirow2, switch minirow1 and minirow2.
                                hold = Row3MiniRowCount(0) : Row3MiniRowCount(0) = Row3MiniRowCount(1) : Row3MiniRowCount(1) = hold
                                hold = StackPermutationTrackerLocal(0) : StackPermutationTrackerLocal(0) = StackPermutationTrackerLocal(1) : StackPermutationTrackerLocal(1) = hold
                            ElseIf Row2CalcMiniRowCountCodeMinimum < 4 AndAlso Row3MiniRowCount(1) > 0 And Row3MiniRowCount(0) = Row3MiniRowCount(1) Then
                                Row3StackPermutationCode = 2
                            End If
                        End If

                        MinLexCandidateSw = True
                        If Row2StackPermutationCode = 0 Then
                            Select Case FirstNonZeroDigitPositionInRow(3)
                                Case < 1    ' If -1 or 0, do nothing.
                                Case 5
                                    If Row3MiniRowCount(0) > 0 Or Row3MiniRowCount(1) > 1 Then MinLexCandidateSw = False
                                Case 4
                                    If Row3MiniRowCount(0) > 0 Or Row3MiniRowCount(1) = 3 Then MinLexCandidateSw = False
                                Case > 5
                                    If Row3MiniRowCount(0) > 0 Or Row3MiniRowCount(1) > 0 Then MinLexCandidateSw = False
                                Case 3
                                    If Row3MiniRowCount(0) > 0 Then MinLexCandidateSw = False
                                Case 2
                                    If Row3MiniRowCount(0) > 1 Then MinLexCandidateSw = False
                                Case 1
                                    If Row3MiniRowCount(0) = 3 Then MinLexCandidateSw = False
                            End Select
                        End If

                        If MinLexCandidateSw Then
                            StackPermutationTrackerCode = StackPermutationTrackerLocal(0) * 100 + StackPermutationTrackerLocal(1) * 10 + StackPermutationTrackerLocal(2)
                            Select Case StackPermutationTrackerCode
                                Case 12
                                Case 21
                                    SwitchStacks12(HoldPuzzle)
                                Case 102
                                    SwitchStacks01(HoldPuzzle)
                                Case 120
                                    Switch3Stacks120(HoldPuzzle)
                                Case 201
                                    Switch3Stacks201(HoldPuzzle)
                                Case 210
                                    SwitchStacks02(HoldPuzzle)
                            End Select

                            ' Positionally (not considering digit values) "right justify" second and third row non-zero digits within MiniRows.
                            ' Also set the ColumnPermutationCode for Band1.
                            ' For each MiniRow (0 to 2 notation).
                            Band1MiniRowColumnPermutationCode(0) = 0
                            Band1MiniRowColumnPermutationCode(1) = 0
                            Band1MiniRowColumnPermutationCode(2) = 0
                            Select Case Row2MiniRowCount(0)                               ' First MiniRow
                                Case 0 ' Row2
                                    Select Case Row3MiniRowCount(0)
                                        Case 0
                                        Case 1
                                            If HoldPuzzle(row3start) > 0 Then
                                                SwitchColumns02(HoldPuzzle)
                                            ElseIf HoldPuzzle(row3start + 1) > 0 Then
                                                SwitchColumns12(HoldPuzzle)
                                            End If
                                        Case 2
                                            If HoldPuzzle(row3start + 2) = 0 Then
                                                SwitchColumns02(HoldPuzzle)
                                            ElseIf HoldPuzzle(row3start + 1) = 0 Then
                                                SwitchColumns01(HoldPuzzle)
                                            End If
                                            Band1MiniRowColumnPermutationCode(0) += 1
                                        Case 3
                                            Band1MiniRowColumnPermutationCode(0) += 3
                                    End Select
                                Case 1 ' Row 2
                                    If HoldPuzzle(row2start) > 0 Then
                                        SwitchColumns02(HoldPuzzle)
                                    ElseIf HoldPuzzle(row2start + 1) > 0 Then
                                        SwitchColumns12(HoldPuzzle)
                                    End If
                                    If HoldPuzzle(row3start) > 0 Then
                                        If HoldPuzzle(row3start + 1) = 0 Then
                                            SwitchColumns01(HoldPuzzle)
                                        Else
                                            Band1MiniRowColumnPermutationCode(0) += 2
                                        End If
                                    End If
                                Case 2 ' Row 2
                                    If HoldPuzzle(row2start + 2) = 0 Then
                                        SwitchColumns02(HoldPuzzle)
                                    ElseIf HoldPuzzle(row2start + 1) = 0 Then
                                        SwitchColumns01(HoldPuzzle)
                                    End If
                                    Band1MiniRowColumnPermutationCode(0) += 1
                                Case 3 ' Row 2
                                    Band1MiniRowColumnPermutationCode(0) += 3
                            End Select

                            Select Case Row2MiniRowCount(1)                               ' Second MiniRow
                                Case 0 ' Row2
                                    Select Case Row3MiniRowCount(1)
                                        Case 1
                                            If HoldPuzzle(row3start + 3) > 0 Then
                                                SwitchColumns35(HoldPuzzle)
                                            ElseIf HoldPuzzle(row3start + 4) > 0 Then
                                                SwitchColumns45(HoldPuzzle)
                                            End If
                                        Case 2
                                            If HoldPuzzle(row3start + 5) = 0 Then
                                                SwitchColumns35(HoldPuzzle)
                                            ElseIf HoldPuzzle(row3start + 4) = 0 Then
                                                SwitchColumns34(HoldPuzzle)
                                            End If
                                            Band1MiniRowColumnPermutationCode(1) += 1
                                        Case 3
                                            Band1MiniRowColumnPermutationCode(1) += 3
                                    End Select
                                Case 1 ' Row2
                                    If HoldPuzzle(row2start + 3) > 0 Then
                                        SwitchColumns35(HoldPuzzle)
                                    ElseIf HoldPuzzle(row2start + 4) > 0 Then
                                        SwitchColumns45(HoldPuzzle)
                                    End If
                                    If HoldPuzzle(row3start + 3) > 0 Then
                                        If HoldPuzzle(row3start + 4) = 0 Then
                                            SwitchColumns34(HoldPuzzle)
                                        Else
                                            Band1MiniRowColumnPermutationCode(1) += 2
                                        End If
                                    End If
                                Case 2 ' Row2
                                    If HoldPuzzle(row2start + 5) = 0 Then
                                        SwitchColumns35(HoldPuzzle)
                                    ElseIf HoldPuzzle(row2start + 4) = 0 Then
                                        SwitchColumns34(HoldPuzzle)
                                    End If
                                    Band1MiniRowColumnPermutationCode(1) += 1
                                Case 3 ' Row2
                                    Band1MiniRowColumnPermutationCode(1) += 3
                            End Select
                            Select Case Row2MiniRowCount(2)                               ' Third MiniRow
'   ' Case 0 ' Row2                                           ' Unreachable - case zero cannot happen for a valid puzzle.
'   '    Select Case Row3MiniRowCount(2)
'   '        Case 1
'   '            If HoldPuzzle(row3start + 6) > 0 Then
'   '                SwitchColumns68(HoldPuzzle)
'   '            ElseIf HoldPuzzle(row3start + 7) > 0 Then
'   '                SwitchColumns78(HoldPuzzle)
'   '            End If
'   '        Case 2
'   '            If HoldPuzzle(row3start + 8) = 0 Then
'   '                SwitchColumns68(HoldPuzzle)
'   '            ElseIf HoldPuzzle(row3start + 7) = 0 Then
'   '                SwitchColumns78(HoldPuzzle)
'   '            End If
'   '            Band1MiniRowColumnPermutationCode(2) += 1
'   '        Case 3
'   '            Band1MiniRowColumnPermutationCode(2) += 3
'   '    End Select
                                Case 1 ' Row2
                                    If HoldPuzzle(row2start + 6) > 0 Then
                                        SwitchColumns68(HoldPuzzle)
                                    ElseIf HoldPuzzle(row2start + 7) > 0 Then
                                        SwitchColumns78(HoldPuzzle)
                                    End If
                                    If HoldPuzzle(row3start + 6) > 0 Then
                                        If HoldPuzzle(row3start + 7) = 0 Then
                                            SwitchColumns67(HoldPuzzle)
                                        Else
                                            Band1MiniRowColumnPermutationCode(2) += 2
                                        End If
                                    End If
                                Case 2 ' Row2
                                    If HoldPuzzle(row2start + 8) = 0 Then
                                        SwitchColumns68(HoldPuzzle)
                                    ElseIf HoldPuzzle(row2start + 7) = 0 Then
                                        SwitchColumns67(HoldPuzzle)
                                    End If
                                    Band1MiniRowColumnPermutationCode(2) += 1
                                Case 3 ' Row2
                                    Band1MiniRowColumnPermutationCode(2) += 3
                            End Select

                            'Array.ConstrainedCopy(HoldPuzzle, (Row2Candidate - 1) * 9, LocalRows2and3, 0, 9)
                            zstart = (Row2Candidate - 1) * 9
                            For z = 0 To 8 : LocalRows2and3(z) = HoldPuzzle(z + zstart) : Next z
                            'Array.ConstrainedCopy(HoldPuzzle, (Row3Candidate - 1) * 9, LocalRows2and3, 9, 9)
                            zstart = (Row3Candidate - 1) * 9
                            For z = 0 To 8 : LocalRows2and3(z + 9) = HoldPuzzle(z + zstart) : Next z
                            'Array.ConstrainedCopy(ColumnTrackerInit, 0, LocalColumnPermutationTracker, 0, 9)
                            For z = 0 To 8 : LocalColumnPermutationTracker(z) = ColumnTrackerInit(z) : Next z
                            FirstColumnPermutationTrackerIsIdentitySw = True
                            For stackpermutationix = 0 To StackPermutations(Row3StackPermutationCode)
                                If stackpermutationix > 0 Then
                                    FirstColumnPermutationTrackerIsIdentitySw = False
                                    stackx = PermutationStackX(Row3StackPermutationCode)(stackpermutationix)    ' Band1 (2 row) stack switch.
                                    stacky = PermutationStackY(Row3StackPermutationCode)(stackpermutationix)
                                    switchx = 3 * stackx
                                    switchy = 3 * stacky
                                    hold = Band1MiniRowColumnPermutationCode(stackx)
                                    Band1MiniRowColumnPermutationCode(stackx) = Band1MiniRowColumnPermutationCode(stacky)
                                    Band1MiniRowColumnPermutationCode(stacky) = hold
                                    hold = Row3MiniRowCount(stackx)
                                    Row3MiniRowCount(stackx) = Row3MiniRowCount(stacky)
                                    Row3MiniRowCount(stacky) = hold
                                    hold = LocalColumnPermutationTracker(switchx)
                                    LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                    LocalColumnPermutationTracker(switchy) = hold
                                    hold = LocalRows2and3(switchx)
                                    LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                    LocalRows2and3(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalColumnPermutationTracker(switchx)
                                    LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                    LocalColumnPermutationTracker(switchy) = hold
                                    hold = LocalRows2and3(switchx)
                                    LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                    LocalRows2and3(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalColumnPermutationTracker(switchx)
                                    LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                    LocalColumnPermutationTracker(switchy) = hold
                                    hold = LocalRows2and3(switchx)
                                    LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                    LocalRows2and3(switchy) = hold
                                    switchx += 7
                                    switchy += 7
                                    hold = LocalRows2and3(switchx)
                                    LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                    LocalRows2and3(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalRows2and3(switchx)
                                    LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                    LocalRows2and3(switchy) = hold
                                    switchx += 1
                                    switchy += 1
                                    hold = LocalRows2and3(switchx)
                                    LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                    LocalRows2and3(switchy) = hold
                                End If

                                MinLexCandidateSw = True
                                Select Case FirstNonZeroDigitPositionInRow(3)
                                    Case < 1    ' If -1 or 0, do nothing.
                                    Case 5
                                        If Row3MiniRowCount(0) > 0 Or LocalRows2and3(12) > 0 Or LocalRows2and3(13) > 0 Then MinLexCandidateSw = False
                                    Case 4
                                        If Row3MiniRowCount(0) > 0 Or LocalRows2and3(12) > 0 Then MinLexCandidateSw = False
                                    Case > 5
                                        If Row3MiniRowCount(0) > 0 Or Row3MiniRowCount(1) > 0 Then MinLexCandidateSw = False
                                    Case 3
                                        If Row3MiniRowCount(0) > 0 Then MinLexCandidateSw = False
                                    Case 2
                                        If LocalRows2and3(9) > 0 Or LocalRows2and3(10) > 0 Then MinLexCandidateSw = False
                                    Case 1
                                        If LocalRows2and3(9) > 0 Then MinLexCandidateSw = False
                                End Select

                                If MinLexCandidateSw Then
                                    ColumnPermutationCode = Band1MiniRowColumnPermutationCode(0) * 16 + Band1MiniRowColumnPermutationCode(1) * 4 + Band1MiniRowColumnPermutationCode(2)
                                    For columnpermutationix = 0 To ColumnPermutations(ColumnPermutationCode)
                                        If columnpermutationix > 0 Then                    ' Band1 (2 row) column permutation.
                                            FirstColumnPermutationTrackerIsIdentitySw = False
                                            switchx = PermutationColumnX(ColumnPermutationCode)(columnpermutationix)
                                            switchy = PermutationColumnY(ColumnPermutationCode)(columnpermutationix)
                                            hold = LocalColumnPermutationTracker(switchx)
                                            LocalColumnPermutationTracker(switchx) = LocalColumnPermutationTracker(switchy)
                                            LocalColumnPermutationTracker(switchy) = hold
                                            hold = LocalRows2and3(switchx)                     ' row 2
                                            LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                            LocalRows2and3(switchy) = hold
                                            switchx += 9
                                            switchy += 9
                                            hold = LocalRows2and3(switchx)                     ' row 3
                                            LocalRows2and3(switchx) = LocalRows2and3(switchy)
                                            LocalRows2and3(switchy) = hold
                                        Else
                                            For i = 9 To 17
                                                If LocalRows2and3(i) > 0 Then
                                                    Row3TestFirstNonZeroDigitPositionInRow = i - 9                      ' Note: 0-8 notation.
                                                    If FirstNonZeroDigitPositionInRow(3) < Row3TestFirstNonZeroDigitPositionInRow Then
                                                        FirstNonZeroDigitPositionInRow(3) = Row3TestFirstNonZeroDigitPositionInRow
                                                    End If
                                                    Exit For
                                                End If
                                            Next i
                                        End If

                                        If FirstNonZeroDigitPositionInRow(3) = Row3TestFirstNonZeroDigitPositionInRow Then
                                            'Array.Clear(DigitAlreadyHitSw, 0, 10)
                                            For z = 0 To 9 : DigitAlreadyHitSw(z) = False : Next z
                                            'Array.Clear(TestRowsRelabeled, 0, FirstNonZeroPositionInRow + 1)
                                            For z = 0 To FirstNonZeroPositionInRow2 : TestRowsRelabeled(z) = 0 : Next z
                                            'Array.ConstrainedCopy(DigitsRelabelWrkInit, 0, DigitsRelabelWrk, 0, 10)      ' Initialize DigitsRelabelWrk 1 to 9 to 10's (the zero element = 0).
                                            For z = 0 To 9 : DigitsRelabelWrk(z) = DigitsRelabelWrkInit(z) : Next z
                                            RelabelLastDigit = 0
                                            For i = FirstNonZeroPositionInRow2 To 17                                       ' Build DigitsRelabelWrk and TestGridRelabeled for
                                                If LocalRows2and3(i) > 0 AndAlso Not DigitAlreadyHitSw(LocalRows2and3(i)) Then
                                                    DigitAlreadyHitSw(LocalRows2and3(i)) = True
                                                    RelabelLastDigit += 1
                                                    DigitsRelabelWrk(LocalRows2and3(i)) = RelabelLastDigit
                                                End If
                                                TestRowsRelabeled(i) = DigitsRelabelWrk(LocalRows2and3(i))
                                            Next i
                                            For i = FirstNonZeroPositionInRow2 To 17
                                                If TestRowsRelabeled(i) > MinLexRows(i) Then   ' Check if row2 is a candidate.
                                                    MinLexCandidateSw = False
                                                    Exit For
                                                ElseIf TestRowsRelabeled(i) < MinLexRows(i) Then
                                                    Exit For
                                                End If
                                            Next i
                                            If MinLexCandidateSw Then
                                                If i < 18 Then
                                                    'Array.ConstrainedCopy(TestRowsRelabeled, 0, MinLexRows, 0, 18)
                                                    For z = 0 To 17 : MinLexRows(z) = TestRowsRelabeled(z) : Next z
                                                    MinLexRowHitSw = False
                                                    ResetMinLexRowSw = True
                                                    ColumnPermutationTrackerIx = 0
                                                    ColumnPermutationTrackerIsIdentitySw(0) = FirstColumnPermutationTrackerIsIdentitySw
                                                    'Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, 0, 9)
                                                    For z = 0 To 8 : ColumnPermutationTracker(z) = LocalColumnPermutationTracker(z) : Next z
                                                    'Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, 0, 10)
                                                    For z = 0 To 9 : HoldDigitsRelabelWrk(z) = DigitsRelabelWrk(z) : Next z
                                                Else
                                                    MinLexRowHitSw = True
                                                    ColumnPermutationTrackerIx += 1
                                                    If FoundColumnPermutationTrackerCountMax < ColumnPermutationTrackerIx Then
                                                        FoundColumnPermutationTrackerCountMax = ColumnPermutationTrackerIx + 1
                                                    End If
'                                                   If ColumnPermutationTrackerArrayMax < ColumnPermutationTrackerIx Then
'                                                       ErrorInputRecordStr = Nothing
'                                                       For j = 0 To 80
'                                                           ErrorInputRecordStr += CStr(InputGridBufferChr(inputgridstart + j))
'                                                       Next j
'                                                       Return 3    ' Error, Too many column permutations for tracker array.
'                                                   Else
                                                        ColumnPermutationTrackerIsIdentitySw(ColumnPermutationTrackerIx) = FirstColumnPermutationTrackerIsIdentitySw
                                                        'Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, ColumnPermutationTrackerIx * 9, 9)
                                                        zstart = ColumnPermutationTrackerIx * 9
                                                        For z = 0 To 8 : ColumnPermutationTracker(z + zstart) = LocalColumnPermutationTracker(z) : Next z
                                                        'Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, 0, 10)
                                                        zstart = ColumnPermutationTrackerIx * 10
                                                        For z = 0 To 9 : HoldDigitsRelabelWrk(z + zstart) = DigitsRelabelWrk(z) : Next z
'                                                   End If
                                                End If
                                            End If ' If MinLexCandidateSw
                                        End If ' If FirstNonZeroDigitPositionInRow(3) = Row3TestFirstNonZeroDigitPositionInRow
                                        MinLexCandidateSw = True
                                    Next columnpermutationix
                                End If ' If MinLexCandidateSw
                            Next stackpermutationix
                        End If ' If MinLexCandidateSw

                        If ResetMinLexRowSw Then
                            Step2Row1CandidateIx = 0
                            Step2Row1Candidate(0) = Row1Candidate
                            Step2Row2Candidate(0) = Row2Candidate
                            Step2ColumnPermutationTrackerStartIx(0) = -1
                            Step2ColumnPermutationTrackerCount(0) = ColumnPermutationTrackerIx + 1
                            Array.ConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, 0, 81)
                            'Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, 0, 10)
                            For z = 0 To 9 : HoldDigitAlreadyHitSw(z) = DigitAlreadyHitSw(z) : Next z
                            HoldRelabelLastDigit(0) = RelabelLastDigit
                        ElseIf MinLexRowHitSw Then
                            Step2Row1CandidateIx += 1
                            Step2Row1Candidate(Step2Row1CandidateIx) = Row1Candidate
                            Step2Row2Candidate(Step2Row1CandidateIx) = Row2Candidate
                            Step2ColumnPermutationTrackerStartIx(Step2Row1CandidateIx) = CandidateColumnPermutationTrackerStartIx
                            Step2ColumnPermutationTrackerCount(Step2Row1CandidateIx) = ColumnPermutationTrackerIx - CandidateColumnPermutationTrackerStartIx
                            Array.ConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, Step2Row1CandidateIx * 81, 81)
                            'Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, Step2Row1CandidateIx * 10, 10)
                            zstart = Step2Row1CandidateIx * 10
                            For z = 0 To 9 : HoldDigitAlreadyHitSw(z + zstart) = DigitAlreadyHitSw(z) : Next z
                            HoldRelabelLastDigit(Step2Row1CandidateIx) = RelabelLastDigit
                        End If
                    Next candidate
                    'Array.ConstrainedCopy(MinLexRows, 0, MinLexGridLocal, 9, 18)   ' For cases of all zeros row1, zero out the first row of MinLexGridLocal and copy MinLexRows to the second and third row of MinLexGridLocal.
                    For z = 0 To 8 : MinLexGridLocal(z) = 0 : Next z
                    For z = 0 To 17 : MinLexGridLocal(z + 9) = MinLexRows(z) : Next z
                    ' End MinLexRows2and3 (Row1 empty case)
                End If ' If CalcMiniRowCountCodeMinimum > 0 Then

                CheckThisPassSw = True
                FirstNonZeroRowCandidateIx = 0
                Array.ConstrainedCopy(MinLexGridLocal, 0, TestGridRelabeled, 0, 27)   ' Copy MinLexGridLocal Band1 to TestGridRelebeled.
                Do While FirstNonZeroRowCandidateIx <= Step2Row1CandidateIx
                    Row1Candidate = Step2Row1Candidate(FirstNonZeroRowCandidateIx)
                    Row2Candidate = Step2Row2Candidate(FirstNonZeroRowCandidateIx)
                    Array.ConstrainedCopy(HoldBand1CandidateJustifiedPuzzles, FirstNonZeroRowCandidateIx * 81, TestBand1, 0, 81)
                    'Array.ConstrainedCopy(HoldDigitAlreadyHitSw, FirstNonZeroRowCandidateIx * 10, Row3DigitAlreadyHitSw, 0, 10)
                    zstart = FirstNonZeroRowCandidateIx * 10
                    For z = 0 To 9 : Row3DigitAlreadyHitSw(z) = HoldDigitAlreadyHitSw(z + zstart) : Next z
                    Row3RelabelLastDigit = HoldRelabelLastDigit(FirstNonZeroRowCandidateIx)

                    ' Move next row1, row2 and row3 candidates to rows 1, 2 and 3 respectively.
                    Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 0, 81)
                    Select Case Row1Candidate
                        Case 1                                  ' If next row1 candidate is row 1 and the row2 candidate is row 2 then do nothing.
                            If Row2Candidate = 3 Then           ' If next row1 candidate is row 1 and the row2 candidate is row 3 then switch row 2 and row 3.
                                'Array.ConstrainedCopy(TestBand1, 9, HoldGrid, 18, 9)
                                For z = 9 To 17 : HoldGrid(z + 9) = TestBand1(z) : Next z
                                'Array.ConstrainedCopy(TestBand1, 18, HoldGrid, 9, 9)
                                For z = 9 To 17 : HoldGrid(z) = TestBand1(z + 9) : Next z
                            End If
                        Case 2                                  ' If next row1 candidate is row 2, ...
                            If Row2Candidate = 1 Then           ' if the row2 candidate is row 1 then switch rows 1 and 2,
                                'Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 9, 9)
                                For z = 0 To 8 : HoldGrid(z + 9) = TestBand1(z) : Next z
                                'Array.ConstrainedCopy(TestBand1, 9, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 9) : Next z
                            Else                                ' else the row2 candidate is row 3 so move rows 1, 2 and 3 to rows 3, 1 and 2.
                                'Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 18, 9)
                                For z = 0 To 8 : HoldGrid(z + 18) = TestBand1(z) : Next z
                                'Array.ConstrainedCopy(TestBand1, 9, HoldGrid, 0, 18)
                                For z = 0 To 17 : HoldGrid(z) = TestBand1(z + 9) : Next z
                            End If
                        Case 3                                  ' If next row1 candidate is row 3, ...
                            If Row2Candidate = 1 Then           ' if the row2 candidate is row 1 then move rows 1, 2 and 3 to rows 2, 3 and 1,
                                '    Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 9, 18)
                                For z = 0 To 17 : HoldGrid(z + 9) = TestBand1(z) : Next z
                                'Array.ConstrainedCopy(TestBand1, 18, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 18) : Next z
                            Else                                ' else the row2 candidate is row 2 so switch rows 1 and 3.
                                'Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 18, 9)
                                For z = 0 To 8 : HoldGrid(z + 18) = TestBand1(z) : Next z
                                'Array.ConstrainedCopy(TestBand1, 18, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 18) : Next z
                            End If
                        Case 4                                  ' If next row1 candidate is row 4 move band 1 to band 2, then ...
                            Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 27, 27)
                            If Row2Candidate = 5 Then           ' if the row2 candidate is row 5 then move band 2 to band 1,
                                Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 0, 27)
                            Else                                ' else the row2 candidate is row 6 so move rows 4, 5 and 6 to rows 1, 3, and 2.
                                'Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 27) : Next z
                                'Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 18) : Next z
                                'Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 9, 9)
                                For z = 9 To 17 : HoldGrid(z) = TestBand1(z + 36) : Next z
                            End If
                        Case 5                                  ' If next row1 candidate is row 5 move band 1 to band 2, then ...
                            Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 27, 27)
                            If Row2Candidate = 4 Then           ' if the row2 candidate is row 4 then move rows 4, 5 and 6 to rows 2, 1 and 3,
                                'Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 9, 9)
                                For z = 9 To 17 : HoldGrid(z) = TestBand1(z + 18) : Next z
                                'Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 36) : Next z
                                'Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 27) : Next z
                            Else                                ' else the row2 candidate is row 6 so move rows 4, 5 and 6 to rows 3, 1 and 2.
                                'Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 9) : Next z
                                'Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 0, 18)
                                For z = 0 To 17 : HoldGrid(z) = TestBand1(z + 36) : Next z
                            End If
                        Case 6                                  ' If next row1 candidate is row 6 move band 1 to band 2, then ...
                            Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 27, 27)
                            If Row2Candidate = 4 Then           ' if the row2 candidate is row 4 then move rows 4, 5 and 6 to rows 2, 3 and 1,
                                'Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 9, 18)
                                For z = 9 To 26 : HoldGrid(z) = TestBand1(z + 18) : Next z
                                'Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 45) : Next z
                            Else                                ' else the row2 candidate is row 5 so move rows 4, 5 and 6 to rows 3, 2 and 1.
                                'Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 9) : Next z
                                'Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 9, 9)
                                For z = 9 To 17 : HoldGrid(z) = TestBand1(z + 27) : Next z
                                'Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 45) : Next z
                            End If
                        Case 7                                  ' If next row1 candidate is row 7 move band 1 to band 3, then ...
                            Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 54, 27)
                            If Row2Candidate = 8 Then           ' if the row2 candidate is row 8 then move band 3 to band 1,
                                Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 0, 27)
                            Else                                ' else the row2 candidate is row 9 so move rows 7, 8 and 9 to rows 1, 3, and 2.
                                'Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 54) : Next z
                                'Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 45) : Next z
                                'Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 9, 9)
                                For z = 9 To 17 : HoldGrid(z) = TestBand1(z + 63) : Next z
                            End If
                        Case 8                                  ' If next row1 candidate is row 8 move band 1 to band 3, then ...
                            Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 54, 27)
                            If Row2Candidate = 7 Then           ' if the row2 candidate is row 7 then move rows 7, 8 and 9 to rows 2, 1 and 3,
                                'Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 9, 9)
                                For z = 9 To 17 : HoldGrid(z) = TestBand1(z + 45) : Next z
                                'Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 63) : Next z
                                'Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 54) : Next z
                            Else                                ' else the row2 candidate is row 9 so move rows 7, 8 and 9 to rows 3, 1 and 2.
                                'Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 36) : Next z
                                'Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 0, 18)
                                For z = 0 To 17 : HoldGrid(z) = TestBand1(z + 63) : Next z
                            End If
                        Case 9                                  ' If next row1 candidate is row 9 move band 1 to band 3, then ...
                            Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 54, 27)
                            If Row2Candidate = 7 Then           ' if the row2 candidate is row 7 then move rows 7, 8 and 9 to rows 2, 3 and 1,
                                'Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 9, 18)
                                For z = 9 To 26 : HoldGrid(z) = TestBand1(z + 45) : Next z
                                'Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 72) : Next z
                            Else                                ' else the row2 candidate is row 8 so move rows 7, 8 and 9 to rows 3, 2 and 1.
                                'Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 18, 9)
                                For z = 18 To 26 : HoldGrid(z) = TestBand1(z + 36) : Next z
                                'Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 9, 9)
                                For z = 9 To 17 : HoldGrid(z) = TestBand1(z + 54) : Next z
                                'Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 0, 9)
                                For z = 0 To 8 : HoldGrid(z) = TestBand1(z + 72) : Next z
                            End If
                    End Select
                    ColumnPermutationTrackerCount = Step2ColumnPermutationTrackerCount(FirstNonZeroRowCandidateIx)
                    FirstColumnPermutationTrackerIsIdentitySw = ColumnPermutationTrackerIsIdentitySw(Step2ColumnPermutationTrackerStartIx(FirstNonZeroRowCandidateIx) + 1)
                    For trackercolumnpermutationix = 0 To ColumnPermutationTrackerCount - 1
                        'Array.ConstrainedCopy(ColumnPermutationTracker, Step2ColumnPermutationTrackerStartIx(FirstNonZeroRowCandidateIx) * 9 + 9 * trackercolumnpermutationix, ColumnPermutationTracker, 0, 9)
                        zstart = (Step2ColumnPermutationTrackerStartIx(FirstNonZeroRowCandidateIx) + 1) * 9 + 9 * trackercolumnpermutationix
                        For z = 0 To 8 : LocalColumnPermutationTracker(z) = ColumnPermutationTracker(z + zstart) : Next z
                        'Array.ConstrainedCopy(HoldDigitsRelabelWrk, Step2DigitsRelabelWrkStartIx(FirstNonZeroRowCandidateIx) * 10 + 10 * trackercolumnpermutationix, Row3DigitsRelabelWrk, 0, 10)
                        zstart = (Step2ColumnPermutationTrackerStartIx(FirstNonZeroRowCandidateIx) + 1) * 10 + 10 * trackercolumnpermutationix
                        For z = 0 To 9 : Row3DigitsRelabelWrk(z) = HoldDigitsRelabelWrk(z + zstart) : Next z
                        Array.ConstrainedCopy(HoldGrid, 0, TestBand1, 0, 81)
                        If trackercolumnpermutationix > 0 OrElse Not FirstColumnPermutationTrackerIsIdentitySw Then
                            If LocalColumnPermutationTracker(0) <> 0 Then                               ' Column 1
                                TestBand1(0) = HoldGrid(LocalColumnPermutationTracker(0))
                                TestBand1(9) = HoldGrid(LocalColumnPermutationTracker(0) + 9)
                                TestBand1(18) = HoldGrid(LocalColumnPermutationTracker(0) + 18)
                                TestBand1(27) = HoldGrid(LocalColumnPermutationTracker(0) + 27)
                                TestBand1(36) = HoldGrid(LocalColumnPermutationTracker(0) + 36)
                                TestBand1(45) = HoldGrid(LocalColumnPermutationTracker(0) + 45)
                                TestBand1(54) = HoldGrid(LocalColumnPermutationTracker(0) + 54)
                                TestBand1(63) = HoldGrid(LocalColumnPermutationTracker(0) + 63)
                                TestBand1(72) = HoldGrid(LocalColumnPermutationTracker(0) + 72)
                            End If
                            If LocalColumnPermutationTracker(1) <> 1 Then                               ' Column 2
                                TestBand1(1) = HoldGrid(LocalColumnPermutationTracker(1))
                                TestBand1(10) = HoldGrid(LocalColumnPermutationTracker(1) + 9)
                                TestBand1(19) = HoldGrid(LocalColumnPermutationTracker(1) + 18)
                                TestBand1(28) = HoldGrid(LocalColumnPermutationTracker(1) + 27)
                                TestBand1(37) = HoldGrid(LocalColumnPermutationTracker(1) + 36)
                                TestBand1(46) = HoldGrid(LocalColumnPermutationTracker(1) + 45)
                                TestBand1(55) = HoldGrid(LocalColumnPermutationTracker(1) + 54)
                                TestBand1(64) = HoldGrid(LocalColumnPermutationTracker(1) + 63)
                                TestBand1(73) = HoldGrid(LocalColumnPermutationTracker(1) + 72)
                            End If
                            If LocalColumnPermutationTracker(2) <> 2 Then                               ' Column 3
                                TestBand1(2) = HoldGrid(LocalColumnPermutationTracker(2))
                                TestBand1(11) = HoldGrid(LocalColumnPermutationTracker(2) + 9)
                                TestBand1(20) = HoldGrid(LocalColumnPermutationTracker(2) + 18)
                                TestBand1(29) = HoldGrid(LocalColumnPermutationTracker(2) + 27)
                                TestBand1(38) = HoldGrid(LocalColumnPermutationTracker(2) + 36)
                                TestBand1(47) = HoldGrid(LocalColumnPermutationTracker(2) + 45)
                                TestBand1(56) = HoldGrid(LocalColumnPermutationTracker(2) + 54)
                                TestBand1(65) = HoldGrid(LocalColumnPermutationTracker(2) + 63)
                                TestBand1(74) = HoldGrid(LocalColumnPermutationTracker(2) + 72)
                            End If
                            If LocalColumnPermutationTracker(3) <> 3 Then                               ' Column 4
                                TestBand1(3) = HoldGrid(LocalColumnPermutationTracker(3))
                                TestBand1(12) = HoldGrid(LocalColumnPermutationTracker(3) + 9)
                                TestBand1(21) = HoldGrid(LocalColumnPermutationTracker(3) + 18)
                                TestBand1(30) = HoldGrid(LocalColumnPermutationTracker(3) + 27)
                                TestBand1(39) = HoldGrid(LocalColumnPermutationTracker(3) + 36)
                                TestBand1(48) = HoldGrid(LocalColumnPermutationTracker(3) + 45)
                                TestBand1(57) = HoldGrid(LocalColumnPermutationTracker(3) + 54)
                                TestBand1(66) = HoldGrid(LocalColumnPermutationTracker(3) + 63)
                                TestBand1(75) = HoldGrid(LocalColumnPermutationTracker(3) + 72)
                            End If
                            If LocalColumnPermutationTracker(4) <> 4 Then                               ' Column 5
                                TestBand1(4) = HoldGrid(LocalColumnPermutationTracker(4))
                                TestBand1(13) = HoldGrid(LocalColumnPermutationTracker(4) + 9)
                                TestBand1(22) = HoldGrid(LocalColumnPermutationTracker(4) + 18)
                                TestBand1(31) = HoldGrid(LocalColumnPermutationTracker(4) + 27)
                                TestBand1(40) = HoldGrid(LocalColumnPermutationTracker(4) + 36)
                                TestBand1(49) = HoldGrid(LocalColumnPermutationTracker(4) + 45)
                                TestBand1(58) = HoldGrid(LocalColumnPermutationTracker(4) + 54)
                                TestBand1(67) = HoldGrid(LocalColumnPermutationTracker(4) + 63)
                                TestBand1(76) = HoldGrid(LocalColumnPermutationTracker(4) + 72)
                            End If
                            If LocalColumnPermutationTracker(5) <> 5 Then                               ' Column 6
                                TestBand1(5) = HoldGrid(LocalColumnPermutationTracker(5))
                                TestBand1(14) = HoldGrid(LocalColumnPermutationTracker(5) + 9)
                                TestBand1(23) = HoldGrid(LocalColumnPermutationTracker(5) + 18)
                                TestBand1(32) = HoldGrid(LocalColumnPermutationTracker(5) + 27)
                                TestBand1(41) = HoldGrid(LocalColumnPermutationTracker(5) + 36)
                                TestBand1(50) = HoldGrid(LocalColumnPermutationTracker(5) + 45)
                                TestBand1(59) = HoldGrid(LocalColumnPermutationTracker(5) + 54)
                                TestBand1(68) = HoldGrid(LocalColumnPermutationTracker(5) + 63)
                                TestBand1(77) = HoldGrid(LocalColumnPermutationTracker(5) + 72)
                            End If
                            If LocalColumnPermutationTracker(6) <> 6 Then                               ' Column 7
                                TestBand1(6) = HoldGrid(LocalColumnPermutationTracker(6))
                                TestBand1(15) = HoldGrid(LocalColumnPermutationTracker(6) + 9)
                                TestBand1(24) = HoldGrid(LocalColumnPermutationTracker(6) + 18)
                                TestBand1(33) = HoldGrid(LocalColumnPermutationTracker(6) + 27)
                                TestBand1(42) = HoldGrid(LocalColumnPermutationTracker(6) + 36)
                                TestBand1(51) = HoldGrid(LocalColumnPermutationTracker(6) + 45)
                                TestBand1(60) = HoldGrid(LocalColumnPermutationTracker(6) + 54)
                                TestBand1(69) = HoldGrid(LocalColumnPermutationTracker(6) + 63)
                                TestBand1(78) = HoldGrid(LocalColumnPermutationTracker(6) + 72)
                            End If
                            If LocalColumnPermutationTracker(7) <> 7 Then                               ' Column 8
                                TestBand1(7) = HoldGrid(LocalColumnPermutationTracker(7))
                                TestBand1(16) = HoldGrid(LocalColumnPermutationTracker(7) + 9)
                                TestBand1(25) = HoldGrid(LocalColumnPermutationTracker(7) + 18)
                                TestBand1(34) = HoldGrid(LocalColumnPermutationTracker(7) + 27)
                                TestBand1(43) = HoldGrid(LocalColumnPermutationTracker(7) + 36)
                                TestBand1(52) = HoldGrid(LocalColumnPermutationTracker(7) + 45)
                                TestBand1(61) = HoldGrid(LocalColumnPermutationTracker(7) + 54)
                                TestBand1(70) = HoldGrid(LocalColumnPermutationTracker(7) + 63)
                                TestBand1(79) = HoldGrid(LocalColumnPermutationTracker(7) + 72)
                            End If
                            If LocalColumnPermutationTracker(8) <> 8 Then                               ' Column 9
                                TestBand1(8) = HoldGrid(LocalColumnPermutationTracker(8))
                                TestBand1(17) = HoldGrid(LocalColumnPermutationTracker(8) + 9)
                                TestBand1(26) = HoldGrid(LocalColumnPermutationTracker(8) + 18)
                                TestBand1(35) = HoldGrid(LocalColumnPermutationTracker(8) + 27)
                                TestBand1(44) = HoldGrid(LocalColumnPermutationTracker(8) + 36)
                                TestBand1(53) = HoldGrid(LocalColumnPermutationTracker(8) + 45)
                                TestBand1(62) = HoldGrid(LocalColumnPermutationTracker(8) + 54)
                                TestBand1(71) = HoldGrid(LocalColumnPermutationTracker(8) + 63)
                                TestBand1(80) = HoldGrid(LocalColumnPermutationTracker(8) + 72)
                            End If
                        End If

                        For i = 0 To 8                                   ' Mark Fixed Columns for the MinLexed Band1 (The Fixed Columns designation will be the same for all Candidates.)
                            FixedColumns(i) = 0
                            If TestBand1(i) > 0 Or TestBand1(i + 9) > 0 Or TestBand1(i + 18) > 0 Then
                                FixedColumns(i) = 1
                            End If
                        Next i
                        If FixedColumns(1) = 1 And FixedColumns(4) = 1 And FixedColumns(7) = 1 Then
                            StillJustifyingSw = False
                            StoppedJustifyingRow = 3
                        Else
                            StillJustifyingSw = True
                        End If

                        MinLexCandidateSw = True

                        '  Identify candidates for Row 4.
                        '  Row 4 Test: After right justification, test each row 4 to 9 with first non-zero digit position.
                        '              And then, if more than one choice compare relabled digit values.
                        Row4TestFirstNonZeroDigitPositionInRow = -1
                        FindFirstNonZeroDigitInRow(4, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(4), FirstNonZeroDigitRelabeled(4))
                        If Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(4) Then
                            Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(4)
                            StartEqualCheck = 4
                        End If
                        FindFirstNonZeroDigitInRow(5, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(5), FirstNonZeroDigitRelabeled(5))
                        If Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(5) Then
                            Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(5)
                            StartEqualCheck = 5
                        End If
                        FindFirstNonZeroDigitInRow(6, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(6), FirstNonZeroDigitRelabeled(6))
                        If Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(6) Then
                            Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(6)
                            StartEqualCheck = 6
                        End If
                        FindFirstNonZeroDigitInRow(7, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(7), FirstNonZeroDigitRelabeled(7))
                        If Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(7) Then
                            Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(7)
                            StartEqualCheck = 7
                        End If
                        FindFirstNonZeroDigitInRow(8, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(8), FirstNonZeroDigitRelabeled(8))
                        If Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(8) Then
                            Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(8)
                            StartEqualCheck = 8
                        End If
                        FindFirstNonZeroDigitInRow(9, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(9), FirstNonZeroDigitRelabeled(9))
                        If Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(9) Then
                            Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(9)
                            StartEqualCheck = 9
                        End If
                        j = -1
                        CandidateFirstRelabeledDigit = 10
                        For i = StartEqualCheck To 9
                            If FirstNonZeroDigitPositionInRow(i) = Row4TestFirstNonZeroDigitPositionInRow Then
                                j += 1
                                Row4TestPositionalCandidateRow(j) = i
                                iForHit(j) = i
                                If CandidateFirstRelabeledDigit > FirstNonZeroDigitRelabeled(i) Then
                                    CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(i)
                                End If
                            End If
                        Next i
                        If j > 0 Then
                            Row4TestCandidateRowIx = -1
                            For i = 0 To j
                                If CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(iForHit(i)) Then
                                    Row4TestCandidateRowIx += 1
                                    Row4TestCandidateRow(Row4TestCandidateRowIx) = Row4TestPositionalCandidateRow(i)
                                End If
                            Next i
                        Else
                            Row4TestCandidateRowIx = 0
                            Row4TestCandidateRow(0) = Row4TestPositionalCandidateRow(0)
                        End If
                        If Not CheckThisPassSw Then
                            If Row4TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow(4) Then
                                Row4TestCandidateRowIx = -1
                            ElseIf CandidateFirstRelabeledDigit < 10 AndAlso (Row4TestFirstNonZeroDigitPositionInRow = MinLexFirstNonZeroDigitPositionInRow(4) And
                                    CandidateFirstRelabeledDigit > MinLexGridLocal(27 + Row4TestFirstNonZeroDigitPositionInRow)) Then
                                Row4TestCandidateRowIx = -1
                            End If
                        End If
                        If Row4TestCandidateRowIx > 0 Then
                            'Array.ConstrainedCopy(FixedColumns, 0, Row3FixedColumns, 0, 9)                 ' Save FixedColumns as of after Row 3.
                            For z = 0 To 8 : Row3FixedColumns(z) = FixedColumns(z) : Next z
                        End If
                        For band2row4orderix = 0 To Row4TestCandidateRowIx                        ' Process each row4 candidate.
                            Array.ConstrainedCopy(TestBand1, 0, TestBand2, 0, 81)
                            Select Case Row4TestCandidateRow(band2row4orderix)                    ' Move the next Row4Test candidate to row 4.
                                Case 4  ' Do nothing.
                                Case 5
                                    'Array.ConstrainedCopy(TestBand1, 36, TestBand2, 27, 9)   ' Move row 5 to 4.
                                    For z = 27 To 35 : TestBand2(z) = TestBand1(z + 9) : Next z
                                    'Array.ConstrainedCopy(TestBand1, 27, TestBand2, 36, 9)   ' Move row 4 to 5.
                                    For z = 27 To 35 : TestBand2(z + 9) = TestBand1(z) : Next z
                                Case 6
                                    'Array.ConstrainedCopy(TestBand1, 45, TestBand2, 27, 9)   ' Move row 6 to 4.
                                    For z = 27 To 35 : TestBand2(z) = TestBand1(z + 18) : Next z
                                    'Array.ConstrainedCopy(TestBand1, 27, TestBand2, 45, 9)   ' Move row 4 to 6.
                                    For z = 27 To 35 : TestBand2(z + 18) = TestBand1(z) : Next z
                                Case 7
                                    Array.ConstrainedCopy(TestBand1, 54, TestBand2, 27, 27)  ' Move band 3 to 2.
                                    Array.ConstrainedCopy(TestBand1, 27, TestBand2, 54, 27)  ' Move band 2 to 3.
                                Case 8
                                    'Array.ConstrainedCopy(TestBand1, 63, TestBand2, 27, 18)  ' Move rows 8 and 9 to 4 and 5.
                                    For z = 27 To 44 : TestBand2(z) = TestBand1(z + 36) : Next z
                                    'Array.ConstrainedCopy(TestBand1, 54, TestBand2, 45, 9)   ' Move row 7 to 6.
                                    For z = 45 To 53 : TestBand2(z) = TestBand1(z + 9) : Next z
                                    Array.ConstrainedCopy(TestBand1, 27, TestBand2, 54, 27)  ' Move band 2 to band 3.
                                Case 9
                                    'Array.ConstrainedCopy(TestBand1, 72, TestBand2, 27, 9)   ' Move row 9 to 4.
                                    For z = 27 To 35 : TestBand2(z) = TestBand1(z + 45) : Next z
                                    'Array.ConstrainedCopy(TestBand1, 54, TestBand2, 36, 18)  ' Move rows 7 and 8 to 5 and 6.
                                    For z = 36 To 53 : TestBand2(z) = TestBand1(z + 18) : Next z
                                    Array.ConstrainedCopy(TestBand1, 27, TestBand2, 54, 27)  ' Move band 2 to band 3.
                            End Select

                            FixedColumnsSavedAsOfRow4Sw = False
                            If StillJustifyingSw Then
                                If band2row4orderix > 0 Then
                                    'Array.ConstrainedCopy(Row3FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row 3 setting.
                                    For z = 0 To 8 : FixedColumns(z) = Row3FixedColumns(z) : Next z
                                End If
                                RightJustifyRow(4, TestBand2, StillJustifyingSw, FixedColumns, Row4StackPermutationCode, Row4ColumnPermutationCode)
                                If Not StillJustifyingSw Then
                                    StoppedJustifyingRow = 4
                                End If
                                If Row4StackPermutationCode > 0 Or Row4ColumnPermutationCode > 0 Then
                                    'Array.ConstrainedCopy(FixedColumns, 0, Row4FixedColumns, 0, 9)               ' Save FixedColumns as of after Row 4.
                                    For z = 0 To 8 : Row4FixedColumns(z) = FixedColumns(z) : Next z
                                    FixedColumnsSavedAsOfRow4Sw = True
                                End If
                            Else
                                Row4StackPermutationCode = 0
                                Row4ColumnPermutationCode = 0
                            End If
                            For row4stackpermutationix = 0 To StackPermutations(Row4StackPermutationCode)
                                If row4stackpermutationix > 0 Then
                                    SwitchStacks(TestBand2, PermutationStackX(Row4StackPermutationCode)(row4stackpermutationix), PermutationStackY(Row4StackPermutationCode)(row4stackpermutationix))
                                End If
                                For row4columnpermutationix = 0 To ColumnPermutations(Row4ColumnPermutationCode)
                                    If row4columnpermutationix > 0 Then
                                        SwitchColumns(TestBand2, PermutationColumnX(Row4ColumnPermutationCode)(row4columnpermutationix), PermutationColumnY(Row4ColumnPermutationCode)(row4columnpermutationix))
                                    End If
                                    'Array.ConstrainedCopy(Row3DigitAlreadyHitSw, 0, Row4DigitAlreadyHitSw, 0, 10)
                                    For z = 0 To 9 : Row4DigitAlreadyHitSw(z) = Row3DigitAlreadyHitSw(z) : Next z
                                    'Array.ConstrainedCopy(Row3DigitsRelabelWrk, 0, Row4DigitsRelabelWrk, 0, 10)
                                    For z = 0 To 9 : Row4DigitsRelabelWrk(z) = Row3DigitsRelabelWrk(z) : Next z
                                    Row4RelabelLastDigit = Row3RelabelLastDigit
                                    For i = 27 To 35                                         ' Build Row4DigitsRelabelWrk and TestGridRelabeled for row4
                                        If TestBand2(i) > 0 AndAlso Not Row4DigitAlreadyHitSw(TestBand2(i)) Then
                                            Row4DigitAlreadyHitSw(TestBand2(i)) = True
                                            Row4RelabelLastDigit += 1
                                            Row4DigitsRelabelWrk(TestBand2(i)) = Row4RelabelLastDigit
                                        End If
                                        TestGridRelabeled(i) = Row4DigitsRelabelWrk(TestBand2(i))
                                    Next i
                                    If Not CheckThisPassSw Then
                                        For i = 27 To 35
                                            If TestGridRelabeled(i) > MinLexGridLocal(i) Then   ' Check if Row4 is greater than MinLex.
                                                MinLexCandidateSw = False
                                                Exit For
                                            ElseIf TestGridRelabeled(i) < MinLexGridLocal(i) Then
                                                CheckThisPassSw = True
                                                Exit For
                                            End If
                                        Next i
                                    End If
                                    If StillJustifyingSw AndAlso (row4stackpermutationix > 0 Or row4columnpermutationix > 0) Then
                                        'Array.ConstrainedCopy(Row4FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row 4 setting.
                                        For z = 0 To 8 : FixedColumns(z) = Row4FixedColumns(z) : Next z
                                    End If
                                    If MinLexCandidateSw Then
                                        '  Identify candidates for Row 5.
                                        '  Row 5 Test: After right justification, test each row 5 and 6 with first non-zero digit position and relabeled digit.
                                        '              And then, if more than one choice compare relabeled digit values.
                                        FindFirstNonZeroDigitInRow(5, StillJustifyingSw, TestBand2, FixedColumns, Row4DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(5), FirstNonZeroDigitRelabeled(5))
                                        FindFirstNonZeroDigitInRow(6, StillJustifyingSw, TestBand2, FixedColumns, Row4DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(6), FirstNonZeroDigitRelabeled(6))

                                        CandidateFirstRelabeledDigit = 99
                                        If FirstNonZeroDigitPositionInRow(5) > FirstNonZeroDigitPositionInRow(6) Then
                                            Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(5)
                                            Row5TestCandidateRowIx = 0
                                            Row5TestCandidateRow(0) = 5
                                        ElseIf FirstNonZeroDigitPositionInRow(5) < FirstNonZeroDigitPositionInRow(6) Then
                                            Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(6)
                                            Row5TestCandidateRowIx = 0
                                            Row5TestCandidateRow(0) = 6
                                        ElseIf FirstNonZeroDigitRelabeled(5) < FirstNonZeroDigitRelabeled(6) Then   ' Positions are equal.
                                            Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(5)
                                            Row5TestCandidateRowIx = 0
                                            Row5TestCandidateRow(0) = 5
                                            CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(5)
                                        ElseIf FirstNonZeroDigitRelabeled(5) > FirstNonZeroDigitRelabeled(6) Then
                                            Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(6)
                                            Row5TestCandidateRowIx = 0
                                            Row5TestCandidateRow(0) = 6
                                            CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(6)
                                        Else
                                            Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(5)
                                            Row5TestCandidateRowIx = 1                  ' This Case: Row 5 and 6 candidates are in the same position within the row and
                                            Row5TestCandidateRow(0) = 5                 ' they both have the same relabeled value. For a valid Sudoku puzzle, that means the
                                            Row5TestCandidateRow(1) = 6                 ' both have the default (unassigned) value of 10 or permutations produced the result, in which case they both need to be checked.
                                            CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(5)
                                        End If
                                        If Not CheckThisPassSw Then
                                            If Row5TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow(5) Then
                                                Row5TestCandidateRowIx = -1
                                            ElseIf CandidateFirstRelabeledDigit < 10 AndAlso (Row5TestFirstNonZeroDigitPositionInRow = MinLexFirstNonZeroDigitPositionInRow(5) And
                                                    CandidateFirstRelabeledDigit > MinLexGridLocal(36 + Row5TestFirstNonZeroDigitPositionInRow)) Then
                                                Row5TestCandidateRowIx = -1
                                            End If
                                        End If

                                        If Not FixedColumnsSavedAsOfRow4Sw And Row5TestCandidateRowIx > 0 Then
                                            'Array.ConstrainedCopy(FixedColumns, 0, Row4FixedColumns, 0, 9)
                                            For z = 0 To 8 : Row4FixedColumns(z) = FixedColumns(z) : Next z
                                            FixedColumnsSavedAsOfRow4Sw = True
                                        End If
                                        For band2rows5and6orderix = 0 To Row5TestCandidateRowIx                    ' Process rows 5 and 6.
                                            If Row5TestCandidateRow(band2rows5and6orderix) = 6 Then
                                                'Array.ConstrainedCopy(TestBand2, 36, HoldRow, 0, 9)      ' Switch rows 5 and 6.
                                                For z = 0 To 8 : HoldRow(z) = TestBand2(z + 36) : Next z
                                                'Array.ConstrainedCopy(TestBand2, 45, TestBand2, 36, 9)
                                                For z = 36 To 44 : TestBand2(z) = TestBand2(z + 9) : Next z
                                                'Array.ConstrainedCopy(HoldRow, 0, TestBand2, 45, 9)
                                                For z = 0 To 8 : TestBand2(z + 45) = HoldRow(z) : Next z
                                            End If
                                            If StillJustifyingSw Then
                                                If band2rows5and6orderix > 0 Then
                                                    'Array.ConstrainedCopy(Row4FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row 4 setting.
                                                    For z = 0 To 8 : FixedColumns(z) = Row4FixedColumns(z) : Next z
                                                End If
                                                RightJustifyRow(5, TestBand2, StillJustifyingSw, FixedColumns, Row5StackPermutationCode, Row5ColumnPermutationCode)
                                                If Not StillJustifyingSw Then
                                                    StoppedJustifyingRow = 5
                                                End If
                                                If Row5StackPermutationCode > 0 Or Row5ColumnPermutationCode > 0 Then
                                                    'Array.ConstrainedCopy(FixedColumns, 0, Row5FixedColumns, 0, 9)               ' Save FixedColumns as of after Row 5.
                                                    For z = 0 To 8 : Row5FixedColumns(z) = FixedColumns(z) : Next z
                                                End If
                                            Else
                                                Row5StackPermutationCode = 0
                                                Row5ColumnPermutationCode = 0
                                            End If

                                            For row5stackpermutationix = 0 To StackPermutations(Row5StackPermutationCode)
                                                If row5stackpermutationix > 0 Then
                                                    SwitchStacks(TestBand2, PermutationStackX(Row5StackPermutationCode)(row5stackpermutationix), PermutationStackY(Row5StackPermutationCode)(row5stackpermutationix))
                                                End If
                                                For row5columnpermutationix = 0 To ColumnPermutations(Row5ColumnPermutationCode)
                                                    If row5columnpermutationix > 0 Then
                                                        SwitchColumns(TestBand2, PermutationColumnX(Row5ColumnPermutationCode)(row5columnpermutationix), PermutationColumnY(Row5ColumnPermutationCode)(row5columnpermutationix))
                                                    End If
                                                    'Array.ConstrainedCopy(Row4DigitAlreadyHitSw, 0, Row5DigitAlreadyHitSw, 0, 10)
                                                    For z = 0 To 9 : Row5DigitAlreadyHitSw(z) = Row4DigitAlreadyHitSw(z) : Next z
                                                    'Array.ConstrainedCopy(Row4DigitsRelabelWrk, 0, Row5DigitsRelabelWrk, 0, 10)
                                                    For z = 0 To 9 : Row5DigitsRelabelWrk(z) = Row4DigitsRelabelWrk(z) : Next z
                                                    Row5RelabelLastDigit = Row4RelabelLastDigit
                                                    For i = 36 To 44                                         ' Build Row5DigitsRelabelWrk and TestGridRelabeled for row5
                                                        If TestBand2(i) > 0 AndAlso Not Row5DigitAlreadyHitSw(TestBand2(i)) Then
                                                            Row5DigitAlreadyHitSw(TestBand2(i)) = True
                                                            Row5RelabelLastDigit += 1
                                                            Row5DigitsRelabelWrk(TestBand2(i)) = Row5RelabelLastDigit
                                                        End If
                                                        TestGridRelabeled(i) = Row5DigitsRelabelWrk(TestBand2(i))
                                                    Next i
                                                    If Not CheckThisPassSw Then
                                                        For i = 36 To 44
                                                            If TestGridRelabeled(i) > MinLexGridLocal(i) Then   ' Check if Row5 is greater than MinLex.
                                                                MinLexCandidateSw = False
                                                                Exit For
                                                            ElseIf TestGridRelabeled(i) < MinLexGridLocal(i) Then
                                                                CheckThisPassSw = True
                                                                Exit For
                                                            End If
                                                        Next i
                                                    End If
                                                    If MinLexCandidateSw Then
                                                        FixedColumnsSavedAsOfRow6Sw = False
                                                        If StillJustifyingSw Then
                                                            If row5stackpermutationix > 0 Or row5columnpermutationix > 0 Then
                                                                'Array.ConstrainedCopy(Row5FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row 5 setting.
                                                                For z = 0 To 8 : FixedColumns(z) = Row5FixedColumns(z) : Next z
                                                            End If
                                                            RightJustifyRow(6, TestBand2, StillJustifyingSw, FixedColumns, Row6StackPermutationCode, Row6ColumnPermutationCode)
                                                            If Not StillJustifyingSw Then
                                                                StoppedJustifyingRow = 6
                                                            End If
                                                            If Row6StackPermutationCode > 0 Or Row6ColumnPermutationCode > 0 Then
                                                                'Array.ConstrainedCopy(FixedColumns, 0, Row6FixedColumns, 0, 9)               ' Save FixedColumns as of after Row 6.
                                                                For z = 0 To 8 : Row6FixedColumns(z) = FixedColumns(z) : Next z
                                                                FixedColumnsSavedAsOfRow6Sw = True
                                                            End If
                                                        Else
                                                            Row6StackPermutationCode = 0
                                                            Row6ColumnPermutationCode = 0
                                                        End If

                                                        For row6stackpermutationix = 0 To StackPermutations(Row6StackPermutationCode)
                                                            If row6stackpermutationix > 0 Then
                                                                SwitchStacks(TestBand2, PermutationStackX(Row6StackPermutationCode)(row6stackpermutationix), PermutationStackY(Row6StackPermutationCode)(row6stackpermutationix))
                                                            End If
                                                            For row6columnpermutationix = 0 To ColumnPermutations(Row6ColumnPermutationCode)
                                                                If row6columnpermutationix > 0 Then
                                                                    SwitchColumns(TestBand2, PermutationColumnX(Row6ColumnPermutationCode)(row6columnpermutationix), PermutationColumnY(Row6ColumnPermutationCode)(row6columnpermutationix))
                                                                End If
                                                                'Array.ConstrainedCopy(Row5DigitAlreadyHitSw, 0, Row6DigitAlreadyHitSw, 0, 10)
                                                                For z = 0 To 9 : Row6DigitAlreadyHitSw(z) = Row5DigitAlreadyHitSw(z) : Next z
                                                                'Array.ConstrainedCopy(Row5DigitsRelabelWrk, 0, Row6DigitsRelabelWrk, 0, 10)
                                                                For z = 0 To 9 : Row6DigitsRelabelWrk(z) = Row5DigitsRelabelWrk(z) : Next z
                                                                Row6RelabelLastDigit = Row5RelabelLastDigit
                                                                For i = 45 To 53                                         ' Build Row6DigitsRelabelWrk and TestGridRelabeled for row6
                                                                    If TestBand2(i) > 0 AndAlso Not Row6DigitAlreadyHitSw(TestBand2(i)) Then
                                                                        Row6DigitAlreadyHitSw(TestBand2(i)) = True
                                                                        Row6RelabelLastDigit += 1
                                                                        Row6DigitsRelabelWrk(TestBand2(i)) = Row6RelabelLastDigit
                                                                    End If
                                                                    TestGridRelabeled(i) = Row6DigitsRelabelWrk(TestBand2(i))
                                                                Next i
                                                                If Not CheckThisPassSw Then
                                                                    For i = 45 To 53
                                                                        If TestGridRelabeled(i) > MinLexGridLocal(i) Then   ' Check if Row6 is greater than MinLex.
                                                                            MinLexCandidateSw = False
                                                                            Exit For
                                                                        ElseIf TestGridRelabeled(i) < MinLexGridLocal(i) Then
                                                                            CheckThisPassSw = True
                                                                            Exit For
                                                                        End If
                                                                    Next i
                                                                End If
                                                                If MinLexCandidateSw Then                                 ' Process Band3
                                                                    ' Identify candidates for Row 7.
                                                                    '        Row 7 Test: After right justification, test each row 7 to 9 first non-zero digit position and relabeled digit.
                                                                    '              And then, if more than one choice compare relabeled digit values.
                                                                    If StillJustifyingSw AndAlso (row6stackpermutationix > 0 Or row6columnpermutationix > 0) Then
                                                                        'Array.ConstrainedCopy(Row6FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row 6 setting.
                                                                        For z = 0 To 8 : FixedColumns(z) = Row6FixedColumns(z) : Next z
                                                                    End If
                                                                    Row7TestFirstNonZeroDigitPositionInRow = -1
                                                                    FindFirstNonZeroDigitInRow(7, StillJustifyingSw, TestBand2, FixedColumns, Row6DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(7), FirstNonZeroDigitRelabeled(7))
                                                                    If Row7TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(7) Then
                                                                        Row7TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(7)
                                                                        StartEqualCheck = 7
                                                                    End If
                                                                    FindFirstNonZeroDigitInRow(8, StillJustifyingSw, TestBand2, FixedColumns, Row6DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(8), FirstNonZeroDigitRelabeled(8))
                                                                    If Row7TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(8) Then
                                                                        Row7TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(8)
                                                                        StartEqualCheck = 8
                                                                    End If
                                                                    FindFirstNonZeroDigitInRow(9, StillJustifyingSw, TestBand2, FixedColumns, Row6DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(9), FirstNonZeroDigitRelabeled(9))
                                                                    If Row7TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow(9) Then
                                                                        Row7TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(9)
                                                                        StartEqualCheck = 9
                                                                    End If
                                                                    j = -1
                                                                    CandidateFirstRelabeledDigit = 10
                                                                    For i = StartEqualCheck To 9
                                                                        If FirstNonZeroDigitPositionInRow(i) = Row7TestFirstNonZeroDigitPositionInRow Then
                                                                            j += 1
                                                                            Row7TestPositionalCandidateRow(j) = i
                                                                            iForHit(j) = i
                                                                            If CandidateFirstRelabeledDigit > FirstNonZeroDigitRelabeled(i) Then
                                                                                CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(i)
                                                                            End If
                                                                        End If
                                                                    Next i
                                                                    If j > 0 Then
                                                                        Row7TestCandidateRowIx = -1
                                                                        For i = 0 To j
                                                                            If CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(iForHit(i)) Then
                                                                                Row7TestCandidateRowIx += 1
                                                                                Row7TestCandidateRow(Row7TestCandidateRowIx) = Row7TestPositionalCandidateRow(i)
                                                                            End If
                                                                        Next i
                                                                    Else
                                                                        Row7TestCandidateRowIx = 0
                                                                        Row7TestCandidateRow(0) = Row7TestPositionalCandidateRow(0)
                                                                    End If
                                                                    If Not CheckThisPassSw Then
                                                                        If Row7TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow(7) Then
                                                                            Row7TestCandidateRowIx = -1
                                                                        ElseIf CandidateFirstRelabeledDigit < 10 AndAlso (Row7TestFirstNonZeroDigitPositionInRow = MinLexFirstNonZeroDigitPositionInRow(7) And
                                                                                    CandidateFirstRelabeledDigit > MinLexGridLocal(54 + Row7TestFirstNonZeroDigitPositionInRow)) Then
                                                                            Row7TestCandidateRowIx = -1
                                                                        End If
                                                                    End If
                                                                    If Not FixedColumnsSavedAsOfRow6Sw And Row7TestCandidateRowIx > 0 Then
                                                                        'Array.ConstrainedCopy(FixedColumns, 0, Row6FixedColumns, 0, 9)               ' Save FixedColumns as of after Row 6.
                                                                        For z = 0 To 8 : Row6FixedColumns(z) = FixedColumns(z) : Next z
                                                                        FixedColumnsSavedAsOfRow6Sw = True
                                                                    End If
                                                                    For band3row7orderix = 0 To Row7TestCandidateRowIx                                ' Process each row 7 candidate.
                                                                        Array.ConstrainedCopy(TestBand2, 0, TestBand3, 0, 81)
                                                                        Select Case Row7TestCandidateRow(band3row7orderix)                            ' Move the next row 7 candidate to row 7.
                                                                            Case 8
                                                                                'Array.ConstrainedCopy(TestBand2, 54, TestBand3, 63, 9)               ' Switch rows 7 and 8.
                                                                                For z = 54 To 62 : TestBand3(z + 9) = TestBand2(z) : Next z
                                                                                'Array.ConstrainedCopy(TestBand2, 63, TestBand3, 54, 9)
                                                                                For z = 54 To 62 : TestBand3(z) = TestBand2(z + 9) : Next z
                                                                            Case 9
                                                                                'Array.ConstrainedCopy(TestBand2, 54, TestBand3, 72, 9)               ' Switch rows 7 and 9.
                                                                                For z = 54 To 62 : TestBand3(z + 18) = TestBand2(z) : Next z
                                                                                'Array.ConstrainedCopy(TestBand2, 72, TestBand3, 54, 9)
                                                                                For z = 54 To 62 : TestBand3(z) = TestBand2(z + 18) : Next z
                                                                        End Select
                                                                        FixedColumnsSavedAsOfRow7Sw = False
                                                                        If StillJustifyingSw Then
                                                                            If row6stackpermutationix > 0 Or row6columnpermutationix > 0 Or band3row7orderix > 0 Then
                                                                                'Array.ConstrainedCopy(Row6FixedColumns, 0, FixedColumns, 0, 9)       ' Reset FixedColumns to after Row 6 setting.
                                                                                For z = 0 To 8 : FixedColumns(z) = Row6FixedColumns(z) : Next z
                                                                            End If
                                                                            RightJustifyRow(7, TestBand3, StillJustifyingSw, FixedColumns, Row7StackPermutationCode, Row7ColumnPermutationCode)
                                                                            If Not StillJustifyingSw Then
                                                                                StoppedJustifyingRow = 7
                                                                            End If
                                                                            If Row7StackPermutationCode > 0 Or Row7ColumnPermutationCode > 0 Then
                                                                                'Array.ConstrainedCopy(FixedColumns, 0, Row7FixedColumns, 0, 9)       ' Save FixedColumns as of after Row 7.
                                                                                For z = 0 To 8 : Row7FixedColumns(z) = FixedColumns(z) : Next z
                                                                                FixedColumnsSavedAsOfRow7Sw = True
                                                                            End If
                                                                        Else
                                                                            Row7StackPermutationCode = 0
                                                                            Row7ColumnPermutationCode = 0
                                                                        End If

                                                                        For row7stackpermutationix = 0 To StackPermutations(Row7StackPermutationCode)
                                                                            If row7stackpermutationix > 0 Then
                                                                                SwitchStacks(TestBand3, PermutationStackX(Row7StackPermutationCode)(row7stackpermutationix), PermutationStackY(Row7StackPermutationCode)(row7stackpermutationix))
                                                                            End If
                                                                            For row7columnpermutationix = 0 To ColumnPermutations(Row7ColumnPermutationCode)
                                                                                If row7columnpermutationix > 0 Then
                                                                                    SwitchColumns(TestBand3, PermutationColumnX(Row7ColumnPermutationCode)(row7columnpermutationix), PermutationColumnY(Row7ColumnPermutationCode)(row7columnpermutationix))
                                                                                End If
                                                                                'Array.ConstrainedCopy(Row6DigitAlreadyHitSw, 0, Row7DigitAlreadyHitSw, 0, 10)
                                                                                For z = 0 To 9 : Row7DigitAlreadyHitSw(z) = Row6DigitAlreadyHitSw(z) : Next z
                                                                                'Array.ConstrainedCopy(Row6DigitsRelabelWrk, 0, Row7DigitsRelabelWrk, 0, 10)
                                                                                For z = 0 To 9 : Row7DigitsRelabelWrk(z) = Row6DigitsRelabelWrk(z) : Next z
                                                                                Row7RelabelLastDigit = Row6RelabelLastDigit
                                                                                For i = 54 To 62                                         ' Build Row7DigitsRelabelWrk and TestGridRelabeled for row7
                                                                                    If TestBand3(i) > 0 AndAlso Not Row7DigitAlreadyHitSw(TestBand3(i)) Then
                                                                                        Row7DigitAlreadyHitSw(TestBand3(i)) = True
                                                                                        Row7RelabelLastDigit += 1
                                                                                        Row7DigitsRelabelWrk(TestBand3(i)) = Row7RelabelLastDigit
                                                                                    End If
                                                                                    TestGridRelabeled(i) = Row7DigitsRelabelWrk(TestBand3(i))
                                                                                Next i
                                                                                If Not CheckThisPassSw Then
                                                                                    For i = 54 To 62
                                                                                        If TestGridRelabeled(i) > MinLexGridLocal(i) Then   ' Check if Row7 is greater than MinLex.
                                                                                            MinLexCandidateSw = False
                                                                                            Exit For
                                                                                        ElseIf TestGridRelabeled(i) < MinLexGridLocal(i) Then
                                                                                            CheckThisPassSw = True
                                                                                            Exit For
                                                                                        End If
                                                                                    Next i
                                                                                End If
                                                                                If MinLexCandidateSw Then
                                                                                    '  Identify candidates for Row 8.
                                                                                    '  Row 8 Test: After right justification, test each row 8 and 9 with first non-zero digit position.
                                                                                    '              And then, if more than one choice compare relabeled digit values.
                                                                                    If StillJustifyingSw AndAlso (row7stackpermutationix > 0 Or row7columnpermutationix > 0) Then
                                                                                        'Array.ConstrainedCopy(Row7FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row 7 setting.
                                                                                        For z = 0 To 8 : FixedColumns(z) = Row7FixedColumns(z) : Next z
                                                                                    End If
                                                                                    FindFirstNonZeroDigitInRow(8, StillJustifyingSw, TestBand3, FixedColumns, Row7DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(8), FirstNonZeroDigitRelabeled(8))
                                                                                    FindFirstNonZeroDigitInRow(9, StillJustifyingSw, TestBand3, FixedColumns, Row7DigitsRelabelWrk, FirstNonZeroDigitPositionInRow(9), FirstNonZeroDigitRelabeled(9))

                                                                                    CandidateFirstRelabeledDigit = 99
                                                                                    If FirstNonZeroDigitPositionInRow(8) > FirstNonZeroDigitPositionInRow(9) Then
                                                                                        Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(8)
                                                                                        Row8TestCandidateRowIx = 0
                                                                                        Row8TestCandidateRow(0) = 8
                                                                                    ElseIf FirstNonZeroDigitPositionInRow(8) < FirstNonZeroDigitPositionInRow(9) Then
                                                                                        Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(9)
                                                                                        Row8TestCandidateRowIx = 0
                                                                                        Row8TestCandidateRow(0) = 9
                                                                                    ElseIf FirstNonZeroDigitRelabeled(8) < FirstNonZeroDigitRelabeled(9) Then   ' Positions are equal.
                                                                                        Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(8)
                                                                                        Row8TestCandidateRowIx = 0
                                                                                        Row8TestCandidateRow(0) = 8
                                                                                        CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(8)
                                                                                    ElseIf FirstNonZeroDigitRelabeled(8) > FirstNonZeroDigitRelabeled(9) Then
                                                                                        Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(9)
                                                                                        Row8TestCandidateRowIx = 0
                                                                                        Row8TestCandidateRow(0) = 9
                                                                                        CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(9)
                                                                                    Else
                                                                                        Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow(8)
                                                                                        Row8TestCandidateRowIx = 1                  ' This Case: Row 8 and 9 candidates are in the same position within the row and
                                                                                        Row8TestCandidateRow(0) = 8                 ' they both have the same relabeled value. For a valid Sudoku puzzle, that means the
                                                                                        Row8TestCandidateRow(1) = 9                 ' both have the default (unassigned) value of 10 or permutations produced the result, in which case they both need to be checked.
                                                                                        CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled(8)
                                                                                    End If
                                                                                    If Not CheckThisPassSw Then
                                                                                        If Row8TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow(8) Then
                                                                                            Row8TestCandidateRowIx = -1
                                                                                        ElseIf CandidateFirstRelabeledDigit < 10 AndAlso (Row8TestFirstNonZeroDigitPositionInRow = MinLexFirstNonZeroDigitPositionInRow(8) And
                                                                                                CandidateFirstRelabeledDigit > MinLexGridLocal(63 + Row8TestFirstNonZeroDigitPositionInRow)) Then
                                                                                            Row8TestCandidateRowIx = -1
                                                                                        End If
                                                                                    End If

                                                                                    If Not FixedColumnsSavedAsOfRow7Sw And Row8TestCandidateRowIx > 0 Then
                                                                                        'Array.ConstrainedCopy(FixedColumns, 0, Row7FixedColumns, 0, 9)
                                                                                        For z = 0 To 8 : Row7FixedColumns(z) = FixedColumns(z) : Next z
                                                                                        FixedColumnsSavedAsOfRow7Sw = True
                                                                                    End If
                                                                                    For band3rows8and9orderix = 0 To Row8TestCandidateRowIx                   ' Process rows 8 and 9.
                                                                                        If Row8TestCandidateRow(band3rows8and9orderix) = 9 Then
                                                                                            'Array.ConstrainedCopy(TestBand3, 63, HoldRow, 0, 9)     ' Switch rows 8 and 9.
                                                                                            For z = 0 To 8 : HoldRow(z) = TestBand3(z + 63) : Next z
                                                                                            'Array.ConstrainedCopy(TestBand3, 72, TestBand3, 63, 9)
                                                                                            For z = 63 To 71 : TestBand3(z) = TestBand3(z + 9) : Next z
                                                                                            'Array.ConstrainedCopy(HoldRow, 0, TestBand3, 72, 9)
                                                                                            For z = 0 To 8 : TestBand3(z + 72) = HoldRow(z) : Next z
                                                                                        End If
                                                                                        If StillJustifyingSw Then
                                                                                            If band3rows8and9orderix > 0 Then
                                                                                                'Array.ConstrainedCopy(Row7FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row7 setting.
                                                                                                For z = 0 To 8 : FixedColumns(z) = Row7FixedColumns(z) : Next z
                                                                                            End If
                                                                                            RightJustifyRow(8, TestBand3, StillJustifyingSw, FixedColumns, Row8StackPermutationCode, Row8ColumnPermutationCode)
                                                                                            If Not StillJustifyingSw Then
                                                                                                StoppedJustifyingRow = 8
                                                                                            End If
                                                                                            If Row8StackPermutationCode > 0 Or Row8ColumnPermutationCode > 0 Then
                                                                                                'Array.ConstrainedCopy(FixedColumns, 0, Row8FixedColumns, 0, 9)               ' Save FixedColumns as of after Row 8.
                                                                                                For z = 0 To 8 : Row8FixedColumns(z) = FixedColumns(z) : Next z
                                                                                            End If
                                                                                        Else
                                                                                            Row8StackPermutationCode = 0
                                                                                            Row8ColumnPermutationCode = 0
                                                                                        End If

                                                                                        For row8stackpermutationix = 0 To StackPermutations(Row8StackPermutationCode)
                                                                                            If row8stackpermutationix > 0 Then
                                                                                                SwitchStacks(TestBand3, PermutationStackX(Row8StackPermutationCode)(row8stackpermutationix), PermutationStackY(Row8StackPermutationCode)(row8stackpermutationix))
                                                                                            End If
                                                                                            For row8columnpermutationix = 0 To ColumnPermutations(Row8ColumnPermutationCode)
                                                                                                If row8columnpermutationix > 0 Then
                                                                                                    SwitchColumns(TestBand3, PermutationColumnX(Row8ColumnPermutationCode)(row8columnpermutationix), PermutationColumnY(Row8ColumnPermutationCode)(row8columnpermutationix))
                                                                                                End If
                                                                                                'Array.ConstrainedCopy(Row7DigitAlreadyHitSw, 0, Row8DigitAlreadyHitSw, 0, 10)
                                                                                                For z = 0 To 9 : Row8DigitAlreadyHitSw(z) = Row7DigitAlreadyHitSw(z) : Next z
                                                                                                'Array.ConstrainedCopy(Row7DigitsRelabelWrk, 0, Row8DigitsRelabelWrk, 0, 10)
                                                                                                For z = 0 To 9 : Row8DigitsRelabelWrk(z) = Row7DigitsRelabelWrk(z) : Next z
                                                                                                Row8RelabelLastDigit = Row7RelabelLastDigit
                                                                                                For i = 63 To 71                                         ' Build Row8DigitsRelabelWrk and TestGridRelabeled for row8
                                                                                                    If TestBand3(i) > 0 AndAlso Not Row8DigitAlreadyHitSw(TestBand3(i)) Then
                                                                                                        Row8DigitAlreadyHitSw(TestBand3(i)) = True
                                                                                                        Row8RelabelLastDigit += 1
                                                                                                        Row8DigitsRelabelWrk(TestBand3(i)) = Row8RelabelLastDigit
                                                                                                    End If
                                                                                                    TestGridRelabeled(i) = Row8DigitsRelabelWrk(TestBand3(i))
                                                                                                Next i
                                                                                                If Not CheckThisPassSw Then
                                                                                                    For i = 63 To 71
                                                                                                        If TestGridRelabeled(i) > MinLexGridLocal(i) Then   ' Check if Row8 is greater than MinLex.
                                                                                                            MinLexCandidateSw = False
                                                                                                            Exit For
                                                                                                        ElseIf TestGridRelabeled(i) < MinLexGridLocal(i) Then
                                                                                                            CheckThisPassSw = True
                                                                                                            Exit For
                                                                                                        End If
                                                                                                    Next i
                                                                                                End If
                                                                                                If MinLexCandidateSw Then
                                                                                                    If StillJustifyingSw Then
                                                                                                        If row8stackpermutationix > 0 Or row8columnpermutationix > 0 Then
                                                                                                            'Array.ConstrainedCopy(Row8FixedColumns, 0, FixedColumns, 0, 9)                  ' Reset FixedColumns to after Row 8 setting.
                                                                                                            For z = 0 To 8 : FixedColumns(z) = Row8FixedColumns(z) : Next z
                                                                                                        End If
                                                                                                        RightJustifyRow(9, TestBand3, StillJustifyingSw, FixedColumns, Row9StackPermutationCode, Row9ColumnPermutationCode)
                                                                                                        If Not StillJustifyingSw Then
                                                                                                            StoppedJustifyingRow = 9
                                                                                                        End If
                                                                                                    Else
                                                                                                        Row9StackPermutationCode = 0
                                                                                                        Row9ColumnPermutationCode = 0
                                                                                                    End If
                                                                                                    For row9stackpermutationix = 0 To StackPermutations(Row9StackPermutationCode)
                                                                                                        If row9stackpermutationix > 0 Then
                                                                                                            SwitchStacks(TestBand3, PermutationStackX(Row9StackPermutationCode)(row9stackpermutationix), PermutationStackY(Row9StackPermutationCode)(row9stackpermutationix))
                                                                                                        End If
                                                                                                        For row9columnpermutationix = 0 To ColumnPermutations(Row9ColumnPermutationCode)
                                                                                                            If row9columnpermutationix > 0 Then
                                                                                                                SwitchColumns(TestBand3, PermutationColumnX(Row9ColumnPermutationCode)(row9columnpermutationix), PermutationColumnY(Row9ColumnPermutationCode)(row9columnpermutationix))
                                                                                                            End If
                                                                                                            'Array.ConstrainedCopy(Row8DigitAlreadyHitSw, 0, Row9DigitAlreadyHitSw, 0, 10)
                                                                                                            For z = 0 To 9 : Row9DigitAlreadyHitSw(z) = Row8DigitAlreadyHitSw(z) : Next z
                                                                                                            'Array.ConstrainedCopy(Row8DigitsRelabelWrk, 0, Row9DigitsRelabelWrk, 0, 10)
                                                                                                            For z = 0 To 9 : Row9DigitsRelabelWrk(z) = Row8DigitsRelabelWrk(z) : Next z
                                                                                                            Row9RelabelLastDigit = Row8RelabelLastDigit
                                                                                                            For i = 72 To 80                                         ' Build Row9DigitsRelabelWrk and TestGridRelabeled for row9
                                                                                                                If TestBand3(i) > 0 AndAlso Not Row9DigitAlreadyHitSw(TestBand3(i)) Then
                                                                                                                    Row9DigitAlreadyHitSw(TestBand3(i)) = True
                                                                                                                    Row9RelabelLastDigit += 1
                                                                                                                    Row9DigitsRelabelWrk(TestBand3(i)) = Row9RelabelLastDigit
                                                                                                                End If
                                                                                                                TestGridRelabeled(i) = Row9DigitsRelabelWrk(TestBand3(i))
                                                                                                            Next i
                                                                                                            If Not CheckThisPassSw Then
                                                                                                                For i = 72 To 80
                                                                                                                    If TestGridRelabeled(i) > MinLexGridLocal(i) Then   ' Check if Row9 is greater than MinLex.
                                                                                                                        MinLexCandidateSw = False
                                                                                                                        Exit For
                                                                                                                    ElseIf TestGridRelabeled(i) < MinLexGridLocal(i) Then
                                                                                                                        CheckThisPassSw = True
                                                                                                                        Exit For
                                                                                                                    End If
                                                                                                                Next i
                                                                                                                If i = 81 Then                                  ' If i = 81 then the TestBand1 matches the last found MinLex Puzzle candidate. If this happens for the actual MinLex Puzzle,
                                                                                                                    MinLexCandidateSw = False                   ' then the puzzle has multiple transformation paths that produce the minimal, thus the puzzle and its grid are automorphic.
                                                                                                                End If
                                                                                                            End If
                                                                                                            If MinLexCandidateSw Then
                                                                                                                Array.ConstrainedCopy(TestGridRelabeled, 0, MinLexGridLocal, 0, 81)
                                                                                                                CheckThisPassSw = False
                                                                                                                MinLexFirstNonZeroDigitPositionInRow(4) = Row4TestFirstNonZeroDigitPositionInRow
                                                                                                                MinLexFirstNonZeroDigitPositionInRow(5) = Row5TestFirstNonZeroDigitPositionInRow
                                                                                                                MinLexFirstNonZeroDigitPositionInRow(7) = Row7TestFirstNonZeroDigitPositionInRow
                                                                                                                MinLexFirstNonZeroDigitPositionInRow(8) = Row8TestFirstNonZeroDigitPositionInRow
                                                                                                                Array.ConstrainedCopy(TestGridRelabeled, 0, MinLexGridLocal, 0, 81)
                                                                                                            End If
                                                                                                            MinLexCandidateSw = True
                                                                                                        Next row9columnpermutationix
                                                                                                    Next row9stackpermutationix
                                                                                                    If StoppedJustifyingRow = 9 Then StoppedJustifyingRow = 0 : StillJustifyingSw = True
                                                                                                End If
                                                                                                MinLexCandidateSw = True
                                                                                            Next row8columnpermutationix
                                                                                        Next row8stackpermutationix
                                                                                        If StoppedJustifyingRow = 8 Then StoppedJustifyingRow = 0 : StillJustifyingSw = True
                                                                                    Next band3rows8and9orderix
                                                                                End If
                                                                                MinLexCandidateSw = True
                                                                            Next row7columnpermutationix
                                                                        Next row7stackpermutationix
                                                                        If StoppedJustifyingRow = 7 Then StoppedJustifyingRow = 0 : StillJustifyingSw = True
                                                                    Next band3row7orderix
                                                                End If
                                                                MinLexCandidateSw = True
                                                            Next row6columnpermutationix
                                                        Next row6stackpermutationix
                                                        If StoppedJustifyingRow = 6 Then StoppedJustifyingRow = 0 : StillJustifyingSw = True
                                                    End If
                                                    MinLexCandidateSw = True
                                                Next row5columnpermutationix
                                            Next row5stackpermutationix
                                            If StoppedJustifyingRow = 5 Then StoppedJustifyingRow = 0 : StillJustifyingSw = True
                                        Next band2rows5and6orderix
                                    End If
                                    MinLexCandidateSw = True
                                Next row4columnpermutationix
                            Next row4stackpermutationix
                            If StoppedJustifyingRow = 4 Then StoppedJustifyingRow = 0 : StillJustifyingSw = True
                        Next band2row4orderix
                    Next trackercolumnpermutationix
                    FirstNonZeroRowCandidateIx += 1
                Loop '  Do While FirstNonZeroRowCandidateIx <= Step2Row1CandidateIx
                ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! sub grid (puzzles) End !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! full grid start !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            Else     ' Process full grid.
                Array.ConstrainedCopy(MinLexFullGridLocalReset, 0, MinLexGridLocal, 0, 81)
                ' Evaluate then InputGrid to determine if first 12 of Minlex is: 123456789456 or 123456789457
                ' It will be "type456" if these exists a pairing of the same 3 digits in any two minirows of diferent rows of any band, otherwise it will be "type457".
                ' For example if the following pattern is detected:   ------------------------------- where the digits 1, 2 and 4 occur in two different minirows,
                ' then the minlex will be of type456.                 | 2  1  4 | x  x  x | x  x  x |
                '                                                     | x  x  x | 2  4  1 | x  x  x |
                '                                                     | x  x  x | x  x  x | x  x  x |
                '                                                     -------------------------------
                ' It is sufficient to check just two cases, for example minirow1 of row1 with minirow2 of row2 and minirow3 of row2, to detect a type456 Minlex.
                MinLexGridLocal(0) = 1 : MinLexGridLocal(1) = 2 : MinLexGridLocal(2) = 3 : MinLexGridLocal(3) = 4 : MinLexGridLocal(4) = 5 : MinLexGridLocal(5) = 6 : MinLexGridLocal(6) = 7 : MinLexGridLocal(7) = 8 : MinLexGridLocal(8) = 9
                MinLexGridLocal(9) = 4 : MinLexGridLocal(10) = 5 : MinLexGridLocal(11) = 7
                MinLexType456Sw = False
                'Array.Clear(MinLexBandType456Sw)
                For z = 0 To 5 : MinLexBandType456Sw(z) = False : Next z
                For i = 0 To 5
                    j = DigitsInMiniRowBit(i * 3, 0) Xor DigitsInMiniRowBit(i * 3 + 1, 1)
                    k = DigitsInMiniRowBit(i * 3, 0) Xor DigitsInMiniRowBit(i * 3 + 1, 2)
                    If j = 0 Or k = 0 Then
                        MinLexBandType456Sw(i) = True
                        If Not MinLexType456Sw Then
                            MinLexGridLocal(11) = 6
                            MinLexGridLocal(12) = 7
                            MinLexGridLocal(13) = 8
                            MinLexGridLocal(14) = 9
                            MinLexType456Sw = True
                        End If
                    End If
                Next i
                If MinLexType456Sw Then
                    For LoopCount = 1 To 2          ' This loop executes two times, first finding the lexicographic minimum equivalent grid for the input grid - direct pass,
                        '                           ' then continuing with the transposed input grid - transposed pass.
                        For band1order = 0 To 2          ' Move each band to the top in order
                            If MinLexBandType456Sw(band1order) Then
                                Select Case band1order
                                    Case 0
                                        Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81)
                                    Case 1
                                        Array.ConstrainedCopy(InputGrid, 27, TestGridBands, 0, 27)             ' Switch Bands 1 & 2.
                                        Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 27, 27)
                                        Array.ConstrainedCopy(InputGrid, 54, TestGridBands, 54, 27)
                                    Case 2
                                        Array.ConstrainedCopy(InputGrid, 54, TestGridBands, 0, 27)             ' Move Band3 to the top.
                                        Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 27, 54)             ' Move Bands 1 & 2 to 2 & 3.
                                End Select
                                For band1row1order = 0 To 2     ' Move each row in band1 to top
                                    Select Case band1row1order
                                        Case 0
                                            Array.ConstrainedCopy(TestGridBands, 0, TestGridRows, 0, 81)
                                        Case 1
                                            For i = 0 To 8                                                     ' Switch Rows 1 & 2.
                                                TestGridRows(i) = TestGridBands(9 + i)
                                                TestGridRows(9 + i) = TestGridBands(i)
                                            Next i
                                            Array.ConstrainedCopy(TestGridBands, 18, TestGridRows, 18, 63)     ' Copy Rows 3 to 9 to TestGridRows.
                                        Case 2
                                            For i = 0 To 8
                                                TestGridRows(i) = TestGridBands(18 + i)                        ' Move Rows 1, 2, 3 to Rows 2, 3, 1 respectively.
                                                TestGridRows(9 + i) = TestGridBands(i)
                                                TestGridRows(18 + i) = TestGridBands(9 + i)
                                            Next i
                                            Array.ConstrainedCopy(TestGridBands, 27, TestGridRows, 27, 54)     ' Copy Rows 4 to 9 to TestGridRows.
                                    End Select
                                    For stackorder = 0 To 5 ' permute stacks
                                        If stackorder > 0 Then
                                            SwitchStacks(TestGridRows, PermutationStackX(3)(stackorder), PermutationStackY(3)(stackorder))
                                        End If
                                        ' Make sure row2 minirow1 containes the same digits as row1 minirow2.
                                        If Not (TestGridRows(3) = TestGridRows(9) Or TestGridRows(3) = TestGridRows(10) Or TestGridRows(3) = TestGridRows(11)) Then
                                            For i = 9 To 17                            ' Switch rows 2 & 3
                                                HoldCell1 = TestGridRows(i)
                                                TestGridRows(i) = TestGridRows(9 + i)
                                                TestGridRows(9 + i) = HoldCell1
                                            Next i
                                        End If
                                        If TestGridRows(3) = TestGridRows(9) Then               ' Reorder row2 minirow1 to match row1 minirow2.
                                            If TestGridRows(4) = TestGridRows(11) Then
                                                SwitchColumns12(TestGridRows)
                                            End If
                                        ElseIf TestGridRows(4) = TestGridRows(9) Then
                                            If TestGridRows(3) = TestGridRows(10) Then
                                                SwitchColumns01(TestGridRows)
                                            Else
                                                Switch3Columns201(TestGridRows)                   ' Right circular shift: 123 ==> 312 (or 012 ==> 201 in 0-8 notation).
                                            End If
                                        ElseIf TestGridRows(4) = TestGridRows(10) Then
                                            SwitchColumns02(TestGridRows)
                                        Else
                                            Switch3Columns120(TestGridRows)                       ' Left circular shift: 123 ==> 231 (or 012 ==> 120 in 0-8 notation).
                                        End If
                                        If TestGridRows(12) = TestGridRows(6) Then              ' Reorder row1 MiniRow3 to match row2 minirow2.
                                            If TestGridRows(13) = TestGridRows(8) Then
                                                SwitchColumns78(TestGridRows)
                                            End If
                                        ElseIf TestGridRows(13) = TestGridRows(6) Then
                                            If TestGridRows(12) = TestGridRows(7) Then
                                                SwitchColumns67(TestGridRows)
                                            Else
                                                Switch3Columns867(TestGridRows)                   ' Right circular shift: 789 ==> 978 (or 678 ==> 867 in 0-8 notation).
                                            End If
                                        ElseIf TestGridRows(13) = TestGridRows(7) Then
                                            SwitchColumns68(TestGridRows)
                                        Else
                                            Switch3Columns786(TestGridRows)                       ' Left circular shift: 789 ==> 897 (or 678 ==> 786 in 0-8 notation).
                                        End If
                                        Array.ConstrainedCopy(TestGridRows, 0, TestGridCols, 0, 81)
                                        For columnorder = 0 To 5
                                            If columnorder > 0 Then            ' Permute all three minirows in sync.
                                                SwitchColumns(TestGridCols, PermutationStackX(3)(columnorder), PermutationStackY(3)(columnorder))
                                                SwitchColumns(TestGridCols, PermutationStackX(3)(columnorder) + 3, PermutationStackY(3)(columnorder) + 3)
                                                SwitchColumns(TestGridCols, PermutationStackX(3)(columnorder) + 6, PermutationStackY(3)(columnorder) + 6)
                                            End If
                                            For i = 1 To 9                     ' Build DigitsRelabelWrk
                                                DigitsRelabelWrk(TestGridCols(i - 1)) = i
                                            Next i
                                            For i = 0 To 26                    ' ReLabel Band1 of TestGridCols into TestGrid123
                                                TestGrid123(i) = DigitsRelabelWrk(TestGridCols(i))
                                            Next i
                                            MinLexCandidateSw = True
                                            For i = 15 To 26                   ' Note: row1 plus first six digits in row2 are 123456789456789, so start checking at row2 seventh digit.
                                                If TestGrid123(i) > MinLexGridLocal(i) Then
                                                    MinLexCandidateSw = False
                                                    Exit For
                                                ElseIf TestGrid123(i) < MinLexGridLocal(i) Then
                                                    Exit For
                                                End If
                                            Next i
                                            If MinLexCandidateSw Then
                                                For i = 27 To 80               ' ReLabel Band2 & Band3 of TestGridCols into TestGrid123
                                                    TestGrid123(i) = DigitsRelabelWrk(TestGridCols(i))
                                                Next i
                                                If TestGrid123(36) < TestGrid123(27) And TestGrid123(36) < TestGrid123(45) Then      ' If row5 is smallest in band2, switch row4 and row5 and
                                                    For i = 0 To 8
                                                        HoldCell1 = TestGrid123(27 + i)
                                                        TestGrid123(27 + i) = TestGrid123(36 + i)
                                                        TestGrid123(36 + i) = HoldCell1
                                                    Next i
                                                ElseIf TestGrid123(45) < TestGrid123(27) And TestGrid123(45) < TestGrid123(36) Then  ' ElseIf row6 is smallest in band2, switch row4 and row6 and                                                                            ' (else, row6 is smallest in band2)
                                                    For i = 0 To 8
                                                        HoldCell1 = TestGrid123(27 + i)
                                                        TestGrid123(27 + i) = TestGrid123(45 + i)
                                                        TestGrid123(45 + i) = HoldCell1
                                                    Next i
                                                End If
                                                If TestGrid123(45) < TestGrid123(36) Then                                            ' If row6 is less than row5,  then switch row5 and row6.
                                                    For i = 0 To 8
                                                        HoldCell1 = TestGrid123(36 + i)
                                                        TestGrid123(36 + i) = TestGrid123(45 + i)
                                                        TestGrid123(45 + i) = HoldCell1
                                                    Next i
                                                End If
                                                If TestGrid123(63) < TestGrid123(54) And TestGrid123(63) < TestGrid123(72) Then      ' If row8 is smallest in band3, then switch row7 and row8.
                                                    For i = 0 To 8
                                                        HoldCell1 = TestGrid123(54 + i)
                                                        TestGrid123(54 + i) = TestGrid123(63 + i)
                                                        TestGrid123(63 + i) = HoldCell1
                                                    Next i
                                                ElseIf TestGrid123(72) < TestGrid123(54) And TestGrid123(72) < TestGrid123(63) Then  ' If row9 is smallest in band3, then switch row7 and row9.
                                                    For i = 0 To 8
                                                        HoldCell1 = TestGrid123(54 + i)
                                                        TestGrid123(54 + i) = TestGrid123(72 + i)
                                                        TestGrid123(72 + i) = HoldCell1
                                                    Next i
                                                End If
                                                If TestGrid123(72) < TestGrid123(63) Then                                             ' If row9 is less than row8, then switch row8 and row9.
                                                    For i = 0 To 8
                                                        HoldCell1 = TestGrid123(63 + i)
                                                        TestGrid123(63 + i) = TestGrid123(72 + i)
                                                        TestGrid123(72 + i) = HoldCell1
                                                    Next i
                                                End If
                                                If TestGrid123(54) < TestGrid123(27) Then                                              ' If row7 < row4 then switch band2 and band3.
                                                    For i = 0 To 8
                                                        HoldCell1 = TestGrid123(27 + i)
                                                        HoldCell2 = TestGrid123(36 + i)
                                                        HoldCell3 = TestGrid123(45 + i)
                                                        TestGrid123(27 + i) = TestGrid123(54 + i)
                                                        TestGrid123(36 + i) = TestGrid123(63 + i)
                                                        TestGrid123(45 + i) = TestGrid123(72 + i)
                                                        TestGrid123(54 + i) = HoldCell1
                                                        TestGrid123(63 + i) = HoldCell2
                                                        TestGrid123(72 + i) = HoldCell3
                                                    Next i
                                                End If
                                                ' Check if TestGrid123 is smaller then current MinLexGridLocal - Note, row1 plus first two digits in row2 are always 12345678945.
                                                For i = 11 To 80
                                                    If TestGrid123(i) > MinLexGridLocal(i) Then
                                                        MinLexCandidateSw = False
                                                        Exit For
                                                    ElseIf TestGrid123(i) < MinLexGridLocal(i) Then
                                                        Exit For
                                                    End If
                                                Next i
                                                If MinLexCandidateSw Then
                                                    Array.ConstrainedCopy(TestGrid123, 0, MinLexGridLocal, 0, 81)
                                                End If
                                            End If
                                        Next columnorder
                                    Next stackorder
                                Next band1row1order
                            End If ' If MinLexBandType456Sw Then
                        Next band1order
                        ' transpose rows / columns, and continue to second loop.
                        Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81)    ' TestGridBands is just used to do this transposition,
                        For i = 0 To 80                                              ' transpose rows to columns.
                            InputGrid(i \ 9 + (i Mod 9) * 9) = TestGridBands(i)
                        Next i
                        MinLexBandType456Sw(0) = MinLexBandType456Sw(3)
                        MinLexBandType456Sw(1) = MinLexBandType456Sw(4)
                        MinLexBandType456Sw(2) = MinLexBandType456Sw(5)
                    Next LoopCount
                Else    ' Process full grid "457" case
                    For LoopCount = 1 To 2          ' This loop executes two times, first finding the lexicographic minimum equivalent grid for the input grid - direct pass,
                        '                           ' then continuing with the transposed input grid - transposed pass.
                        For band1order = 0 To 2          ' Move each band to the top in order
                            Select Case band1order
                                Case 0
                                    Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81)
                                Case 1
                                    Array.ConstrainedCopy(InputGrid, 27, TestGridBands, 0, 27)             ' Switch Bands 1 & 2.
                                    Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 27, 27)
                                    Array.ConstrainedCopy(InputGrid, 54, TestGridBands, 54, 27)
                                Case 2
                                    Array.ConstrainedCopy(InputGrid, 54, TestGridBands, 0, 27)             ' Move Band3 to the top.
                                    Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 27, 54)             ' Move Bands 1 & 2 to 2 & 3.
                            End Select
                            For band1row1order = 0 To 2     ' Move each row in band1 to top
                                Select Case band1row1order
                                    Case 0
                                        Array.ConstrainedCopy(TestGridBands, 0, TestGridRows, 0, 81)
                                    Case 1
                                        For i = 0 To 8                                                     ' Switch Rows 1 & 2.
                                            TestGridRows(i) = TestGridBands(9 + i)
                                            TestGridRows(9 + i) = TestGridBands(i)
                                        Next i
                                        Array.ConstrainedCopy(TestGridBands, 18, TestGridRows, 18, 63)     ' Copy Rows 3 to 9 to TestGridRows.
                                    Case 2
                                        For i = 0 To 8
                                            TestGridRows(i) = TestGridBands(18 + i)                        ' Move Rows 1, 2, 3 to Rows 2, 3, 1 respectively.
                                            TestGridRows(9 + i) = TestGridBands(i)
                                            TestGridRows(18 + i) = TestGridBands(9 + i)
                                        Next i
                                        Array.ConstrainedCopy(TestGridBands, 27, TestGridRows, 27, 54)     ' Copy Rows 4 to 9 to TestGridRows.
                                End Select
                                For stackorder = 0 To 5 ' permute stacks
                                    If stackorder > 0 Then
                                        SwitchStacks(TestGridRows, PermutationStackX(3)(stackorder), PermutationStackY(3)(stackorder))
                                    End If
                                    ' Make sure row2 minirow1 containes two of the same digits as row1 minirow2.
                                    MatchingDigitCount = 0
                                    If TestGridRows(3) = TestGridRows(9) Or TestGridRows(3) = TestGridRows(10) Or TestGridRows(3) = TestGridRows(11) Then
                                        MatchingDigit1 = TestGridRows(3)
                                        MatchingDigit1Position = 3
                                        MatchingDigitCount = 1
                                        If TestGridRows(4) = TestGridRows(9) Or TestGridRows(4) = TestGridRows(10) Or TestGridRows(4) = TestGridRows(11) Then
                                            MatchingDigit2 = TestGridRows(4)
                                            MatchingDigit2Position = 4
                                            MatchingDigitCount = 2
                                        ElseIf TestGridRows(5) = TestGridRows(9) Or TestGridRows(5) = TestGridRows(10) Or TestGridRows(5) = TestGridRows(11) Then
                                            MatchingDigit2 = TestGridRows(5)
                                            MatchingDigit2Position = 5
                                            MatchingDigitCount = 2
                                        End If
                                    ElseIf TestGridRows(4) = TestGridRows(9) Or TestGridRows(4) = TestGridRows(10) Or TestGridRows(4) = TestGridRows(11) Then
                                        MatchingDigit1 = TestGridRows(4)
                                        MatchingDigit1Position = 4
                                        MatchingDigitCount = 1
                                        If TestGridRows(5) = TestGridRows(9) Or TestGridRows(5) = TestGridRows(10) Or TestGridRows(5) = TestGridRows(11) Then
                                            MatchingDigit2 = TestGridRows(5)
                                            MatchingDigit2Position = 5
                                            MatchingDigitCount = 2
                                        End If
                                    End If
                                    If MatchingDigitCount < 2 Then
                                        For i = 9 To 17                            ' Switch rows 2 & 3
                                            HoldCell1 = TestGridRows(i)
                                            TestGridRows(i) = TestGridRows(9 + i)
                                            TestGridRows(9 + i) = HoldCell1
                                        Next i
                                        If TestGridRows(3) = TestGridRows(9) Or TestGridRows(3) = TestGridRows(10) Or TestGridRows(3) = TestGridRows(11) Then
                                            MatchingDigit1 = TestGridRows(3)
                                            MatchingDigit1Position = 3
                                            If TestGridRows(4) = TestGridRows(9) Or TestGridRows(4) = TestGridRows(10) Or TestGridRows(4) = TestGridRows(11) Then
                                                MatchingDigit2 = TestGridRows(4)
                                                MatchingDigit2Position = 4
                                            ElseIf TestGridRows(5) = TestGridRows(9) Or TestGridRows(5) = TestGridRows(10) Or TestGridRows(5) = TestGridRows(11) Then
                                                MatchingDigit2 = TestGridRows(5)
                                                MatchingDigit2Position = 5
                                            End If
                                        Else
                                            MatchingDigit1 = TestGridRows(4)
                                            MatchingDigit1Position = 4
                                            MatchingDigit2 = TestGridRows(5)
                                            MatchingDigit2Position = 5
                                        End If
                                    End If
                                    If MatchingDigit2Position = 5 Then                              ' Adjust row1 minirow2 so the two matching digits are in positions 3 & 4 (0-8 notation)
                                        If MatchingDigit1Position = 4 Then
                                            SwitchColumns35(TestGridRows)
                                        Else
                                            SwitchColumns45(TestGridRows)
                                        End If
                                    End If
                                    If TestGridRows(3) = TestGridRows(9) Then                       ' Reorder two matching digits in row2 minirow1 to match row1 minirow2
                                        If TestGridRows(4) = TestGridRows(11) Then
                                            SwitchColumns12(TestGridRows)
                                        End If
                                    ElseIf TestGridRows(4) = TestGridRows(9) Then
                                        If TestGridRows(3) = TestGridRows(10) Then
                                            SwitchColumns01(TestGridRows)
                                        Else
                                            Switch3Columns201(TestGridRows)                         ' Right circular shift: 123 ==> 312 (or 012 ==> 201 in 0-8 notation).
                                        End If
                                    ElseIf TestGridRows(4) = TestGridRows(10) Then
                                        SwitchColumns02(TestGridRows)
                                    Else
                                        Switch3Columns120(TestGridRows)                             ' Left circular shift: 123 ==> 231 (or 012 ==> 120 in 0-8 notation).
                                    End If
                                    If TestGridRows(11) = TestGridRows(7) Then                      ' Reorder row1 MiniRow3 position 6 to match the third cell in row2.
                                        SwitchColumns67(TestGridRows)
                                    ElseIf TestGridRows(11) = TestGridRows(8) Then
                                        SwitchColumns68(TestGridRows)
                                    End If
                                    '                                                               ' Reorder row1 MiniRow3 positions 7 and 8 to match the order of the matching digits in row2 minirow2.
                                    If TestGridRows(12) = TestGridRows(8) OrElse
                                      (TestGridRows(13) = TestGridRows(8) And TestGridRows(12) <> TestGridRows(7)) Then
                                        SwitchColumns78(TestGridRows)
                                    End If
                                    Array.ConstrainedCopy(TestGridRows, 0, TestGridCols, 0, 81)
                                    For columnorder = 0 To 1
                                        If columnorder > 0 Then                                     ' Permute matching digits in first two minirows in sync.
                                            SwitchColumns01(TestGridCols)
                                            SwitchColumns34(TestGridCols)
                                        End If
                                        '                                                           ' Reorder row1 MiniRow3 positions 7 and 8 to match the order of the matching digits in row2 minirow2.
                                        If TestGridRows(12) = TestGridRows(8) OrElse
                                          (TestGridRows(13) = TestGridRows(8) And TestGridRows(12) <> TestGridRows(7)) Then
                                            SwitchColumns78(TestGridRows)
                                        End If
                                        For i = 1 To 9                     ' Build DigitsRelabelWrk
                                            DigitsRelabelWrk(TestGridCols(i - 1)) = i
                                        Next i
                                        For i = 0 To 26                    ' ReLabel Band1 of TestGridCols into TestGrid123
                                            TestGrid123(i) = DigitsRelabelWrk(TestGridCols(i))
                                        Next i
                                        MinLexCandidateSw = True
                                        For i = 12 To 26                   ' Note: row1 plus first three digits in row2 are 123456789457, so start checking at row2 fourth digit.
                                            If TestGrid123(i) > MinLexGridLocal(i) Then
                                                MinLexCandidateSw = False
                                                Exit For
                                            ElseIf TestGrid123(i) < MinLexGridLocal(i) Then
                                                Exit For
                                            End If
                                        Next i
                                        If MinLexCandidateSw Then
                                            For i = 27 To 80               ' ReLabel Band2 & Band3 of TestGridCols into TestGrid123
                                                TestGrid123(i) = DigitsRelabelWrk(TestGridCols(i))
                                            Next i
                                            If TestGrid123(36) < TestGrid123(27) And TestGrid123(36) < TestGrid123(45) Then    ' If row5 is smallest in band2, switch row4 and row5 and
                                                For i = 0 To 8
                                                    HoldCell1 = TestGrid123(27 + i)
                                                    TestGrid123(27 + i) = TestGrid123(36 + i)
                                                    TestGrid123(36 + i) = HoldCell1
                                                Next i
                                            ElseIf TestGrid123(45) < TestGrid123(27) And TestGrid123(45) < TestGrid123(36) Then  ' ElseIf row6 is smallest in band2, switch row4 and row6 and                                                                            ' (else, row6 is smallest in band2)
                                                For i = 0 To 8
                                                    HoldCell1 = TestGrid123(27 + i)
                                                    TestGrid123(27 + i) = TestGrid123(45 + i)
                                                    TestGrid123(45 + i) = HoldCell1
                                                Next i
                                            End If
                                            If TestGrid123(45) < TestGrid123(36) Then                                            ' If row6 is less than row5,  then switch row5 and row6.
                                                For i = 0 To 8
                                                    HoldCell1 = TestGrid123(36 + i)
                                                    TestGrid123(36 + i) = TestGrid123(45 + i)
                                                    TestGrid123(45 + i) = HoldCell1
                                                Next i
                                            End If
                                            If TestGrid123(63) < TestGrid123(54) And TestGrid123(63) < TestGrid123(72) Then      ' If row8 is smallest in band3, then switch row7 and row8.
                                                For i = 0 To 8
                                                    HoldCell1 = TestGrid123(54 + i)
                                                    TestGrid123(54 + i) = TestGrid123(63 + i)
                                                    TestGrid123(63 + i) = HoldCell1
                                                Next i
                                            ElseIf TestGrid123(72) < TestGrid123(54) And TestGrid123(72) < TestGrid123(63) Then  ' If row9 is smallest in band3, then switch row7 and row9.
                                                For i = 0 To 8
                                                    HoldCell1 = TestGrid123(54 + i)
                                                    TestGrid123(54 + i) = TestGrid123(72 + i)
                                                    TestGrid123(72 + i) = HoldCell1
                                                Next i
                                            End If
                                            If TestGrid123(72) < TestGrid123(63) Then                                             ' If row9 is less than row8, then switch row8 and row9.
                                                For i = 0 To 8
                                                    HoldCell1 = TestGrid123(63 + i)
                                                    TestGrid123(63 + i) = TestGrid123(72 + i)
                                                    TestGrid123(72 + i) = HoldCell1
                                                Next i
                                            End If
                                            If TestGrid123(54) < TestGrid123(27) Then                                              ' If row7 < row4 then switch band2 and band3.
                                                For i = 0 To 8
                                                    HoldCell1 = TestGrid123(27 + i)
                                                    HoldCell2 = TestGrid123(36 + i)
                                                    HoldCell3 = TestGrid123(45 + i)
                                                    TestGrid123(27 + i) = TestGrid123(54 + i)
                                                    TestGrid123(36 + i) = TestGrid123(63 + i)
                                                    TestGrid123(45 + i) = TestGrid123(72 + i)
                                                    TestGrid123(54 + i) = HoldCell1
                                                    TestGrid123(63 + i) = HoldCell2
                                                    TestGrid123(72 + i) = HoldCell3
                                                Next i
                                            End If
                                            ' Check if TestGrid123 is smaller then current MinLexGridLocal - Note, row1 plus first two digits in row2 are always 12345678945.
                                            For i = 11 To 80
                                                If TestGrid123(i) > MinLexGridLocal(i) Then
                                                    MinLexCandidateSw = False
                                                    Exit For
                                                ElseIf TestGrid123(i) < MinLexGridLocal(i) Then
                                                    Exit For
                                                End If
                                            Next i
                                            If MinLexCandidateSw Then
                                                Array.ConstrainedCopy(TestGrid123, 0, MinLexGridLocal, 0, 81)
                                            End If
                                        End If
                                    Next columnorder
                                Next stackorder
                            Next band1row1order
                        Next band1order
                        ' transpose rows / columns, and continue to second loop.
                        Array.ConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81)    ' TestGridBands is just used to do this transposition,
                        For i = 0 To 80                                              ' transpose rows to columns.
                            InputGrid(i \ 9 + (i Mod 9) * 9) = TestGridBands(i)
                        Next i
                    Next LoopCount
                End If
                ' End ProcessFullGrid
                ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! full grid end !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            End If   ' If CluesCount < 81
'           If ProcessMode = "2" Then
                zstart = ResultsBufferCharacterOffset + inputgridix * 83
'               If PatternModeSw Then
'                   For z = 0 To 80                       ' On output, for pattern mode, convert from integer array to char "." or "X".
'                       If MinLexGridLocal(z) = 1 Then
'                           MinLexBufferChr(zstart + z) = "X"c
'                       Else
'                           MinLexBufferChr(zstart + z) = "."c
'                       End If
'                   Next z
'               Else
                    For z = 0 To 80                       ' On output, convert from integer array back to char ".' or 1 - 9.
                        MinLexBufferChr(zstart + z) = CharPeriodTO9(MinLexGridLocal(z))
                    Next z
'               End If
                MinLexBufferChr(zstart + 81) = CarriageReturnChr
                MinLexBufferChr(zstart + 82) = LineFeedChr
'           Else
'               If PatternModeSw Then
'                   For k = 0 To 80                       ' On output, for pattern mode, convert from integer array to char "." or "X".
'                       If MinLexGridLocal(k) = 0 Then
'                           MinLexGridLocalChr(k) = "."c
'                       Else
'                           MinLexGridLocalChr(k) = "X"c
'                       End If
'                   Next k
'               Else
'                   For k = 0 To 80                       ' On output, convert from integer array back to char ".' or 1 - 9.
'                       MinLexGridLocalChr(k) = CharPeriodTO9(MinLexGridLocal(k))
'                   Next k
'               End If
'               MinlexBufferStr(ResultsBufferOffset + inputgridix) = CStr(MinLexGridLocalChr)
'           End If
        Next inputgridix ' For inputgridix = 0 To InputLineCount - 1

'       If Not ProcessMode = "2" Then
'           Array.Sort(MinlexBufferStr, ResultsBufferOffset, InputBufferSize \ 83)
'       End If
'
        Return 0 ' successful run.
    End Function ' Public Sub MinLex9X9SR1

    Shared Sub RightJustifyRow( Row As Integer,
                                Puzzle As Integer(),
                          ByRef StillJusfifyingSw As Boolean,
                                FixedColumns As Integer(),
                          ByRef StackPermutationCode As Integer,
                          ByRef ColumnPermutationCode As Integer)
        ' This routine is used for rows 4 to 9 to move non-zero digits to the right where possible and determine if stack or column permutations are required.
        ' FixedColumns indicates which columns contain a non-zero digit above the provided row. It is reset based on the right justified digit positions of this row.
        Dim i, hold, rowstart As Integer
        Dim MiniRowCount As Integer() = New Integer(2) {}
        Dim MiniRowPermutationCode As Integer() = New Integer(2) {}

        rowstart = (Row - 1) * 9
        StackPermutationCode = 0
        ' Count non-zero digits in MiniRows.
        For i = 0 To 8
            If Puzzle(rowstart + i) > 0 Then
                MiniRowCount(i \ 3) += 1
            End If
        Next i

        ' "right justify" MiniRows 1 andn 2.
        If FixedColumns(5) = 0 Then                         ' If minirow2 is not fixed, then
            If MiniRowCount(0) > MiniRowCount(1) Then       ' If minirow1 is greater than minirow2, switch Stack1 and Stack2.
                hold = MiniRowCount(0) : MiniRowCount(0) = MiniRowCount(1) : MiniRowCount(1) = hold
                SwitchStacks01(Puzzle)
            End If
        End If
        ' "right justify" columns within MiniRows. Note: The FixedColumn indicators are right justified within MiniRows.
        ' For each MiniRow (0 to 2 notation). Note: the comments below use "abc" notation for digits in the MiniRow.
        If FixedColumns(1) = 0 Then        ' First MiniRow.   If the "b" position is not fixed (then the "a" position would also not be fixed.), then
            If Puzzle(rowstart + 0) > 0 And Puzzle(rowstart + 1) = 0 Then              ' If "a" is non-zero and "b' is zero, then switch column "a" and column "b".
                SwitchColumns01(Puzzle)
            End If
            If FixedColumns(2) = 0 AndAlso                  ' If the "c" position is not fixed (then the "a" and "b" positions would also not be fixed.), then
                (Puzzle(rowstart + 1) > 0 And Puzzle(rowstart + 2) = 0) Then      ' If "b" is non-zero and "c' is zero, then switch column "b" and column "c".
                SwitchColumns12(Puzzle)
                If Puzzle(rowstart + 0) > 0 And Puzzle(rowstart + 1) = 0 Then     ' Repeat first If statement above.
                    SwitchColumns01(Puzzle)
                End If
            End If
        End If
        If FixedColumns(4) = 0 Then        ' Second MiniRow.
            If Puzzle(rowstart + 3) > 0 And Puzzle(rowstart + 4) = 0 Then
                SwitchColumns34(Puzzle)
            End If
            If FixedColumns(5) = 0 AndAlso
                (Puzzle(rowstart + 4) > 0 And Puzzle(rowstart + 5) = 0) Then
                SwitchColumns45(Puzzle)
                If Puzzle(rowstart + 3) > 0 And Puzzle(rowstart + 4) = 0 Then
                    SwitchColumns34(Puzzle)
                End If
            End If
        End If
        If FixedColumns(7) = 0 Then        ' Third MiniRow.
            If Puzzle(rowstart + 6) > 0 And Puzzle(rowstart + 7) = 0 Then
                SwitchColumns67(Puzzle)
            End If
        End If

        ' Set Stack PermutationCode
        If FixedColumns(5) = 0 AndAlso (Puzzle(rowstart + 5)) > 0 And MiniRowCount(0) = MiniRowCount(1) Then
            StackPermutationCode = 2
        End If

        ' Set Column PermutationCode
        MiniRowPermutationCode(0) = 0
        MiniRowPermutationCode(1) = 0
        MiniRowPermutationCode(2) = 0
        If FixedColumns(2) = 0 Then      ' First MiniRow
            If Puzzle(rowstart) > 0 Then
                MiniRowPermutationCode(0) = 3
            ElseIf Puzzle(rowstart + 1) > 0 Then
                MiniRowPermutationCode(0) = 1
            End If
        ElseIf FixedColumns(1) = 0 AndAlso
            (Puzzle(rowstart) > 0 And Puzzle(rowstart + 1) > 0) Then
            MiniRowPermutationCode(0) = 2
        End If
        If FixedColumns(5) = 0 Then      ' Second MiniRow
            If Puzzle(rowstart + 3) > 0 Then
                MiniRowPermutationCode(1) = 3
            ElseIf Puzzle(rowstart + 4) > 0 Then
                MiniRowPermutationCode(1) = 1
            End If
        ElseIf FixedColumns(4) = 0 AndAlso
            (Puzzle(rowstart + 3) > 0 And Puzzle(rowstart + 4) > 0) Then
            MiniRowPermutationCode(1) = 2
        End If
        If FixedColumns(7) = 0 AndAlso   ' Third MiniRow
        (Puzzle(rowstart + 6) > 0 And Puzzle(rowstart + 7) > 0) Then
            MiniRowPermutationCode(2) = 2
        End If
        ColumnPermutationCode = MiniRowPermutationCode(0) * 16 + MiniRowPermutationCode(1) * 4 + MiniRowPermutationCode(2)

        For i = 1 To 7                                   ' Mark Fixed Columns.
            If Puzzle(rowstart + i) > 0 Then
                FixedColumns(i) = 1
            End If
        Next i

        If FixedColumns(1) = 1 And FixedColumns(4) = 1 And FixedColumns(7) = 1 Then
            StillJusfifyingSw = False
        Else
            StillJusfifyingSw = True
        End If
    End Sub

    Shared Sub FindFirstNonZeroDigitInRow( Row As Integer,
                                           StillJustifyingSw As Boolean,
                                           Puzzle As Integer(),
                                           FixedColumns As Integer(),
                                           DigitsRelabelWrk As Integer(),
                                     ByRef FirstNonZeroDigitPositionInRow As Integer,
                                     ByRef FirstNonZeroDigitRelabeled As Integer)
        Dim i, hold As Integer
        Dim hold1, hold2, hold3 As Integer
        Dim MiniRowCount() As Integer = New Integer(2) {}
        Dim LocalRow As Integer() = New Integer(8) {}
        Dim CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx As Integer
        Dim CandidateDigitsConsideringPossibleStackAndColumnPermutations As Integer() = New Integer(5) {}
        Dim FirstNonZeroDigit As Integer

        Array.ConstrainedCopy(Puzzle, (Row - 1) * 9, LocalRow, 0, 9)

        If StillJustifyingSw Then
            For i = 0 To 8                                   ' Count non-zero digits in minirow.
                If LocalRow(i) > 0 Then
                    MiniRowCount(i \ 3) += 1
                End If
            Next i
            ' "right justify" MiniRows in row.                    Note: The FixedColumns indicators are right justified within the row.
            If FixedColumns(5) = 0 Then                         ' If minirow2 is not fixed, then minirow1 is also not fixed.
                If MiniRowCount(0) > MiniRowCount(1) Then       ' If minirow1 count is greater than minirow2, switch minirow1 and minirow2.
                    hold = MiniRowCount(0) : MiniRowCount(0) = MiniRowCount(1) : MiniRowCount(1) = hold
                    hold1 = LocalRow(0)
                    hold2 = LocalRow(1)
                    hold3 = LocalRow(2)
                    LocalRow(0) = LocalRow(3)
                    LocalRow(1) = LocalRow(4)
                    LocalRow(2) = LocalRow(5)
                    LocalRow(3) = hold1
                    LocalRow(4) = hold2
                    LocalRow(5) = hold3
                End If
            End If

            ' "right justify" columns in row within MiniRows. Note: The FixedColumn indicators are right justified within MiniRows.
            ' For each MiniRow (0 to 2 notation). Note: the comments below use "abc" notation for digits in the MiniRow.
            If FixedColumns(1) = 0 Then     ' First MiniRow            ' If the "b" position is not fixed (then the "a" position would also not be fixed.), then
                If LocalRow(0) > 0 And LocalRow(1) = 0 Then         ' If "a" is non-zero and "b' is zero, then switch "a" and "b".
                    LocalRow(1) = LocalRow(0)
                    LocalRow(0) = 0
                End If
                If FixedColumns(2) = 0 Then                ' If "c" position is not fixed (then the "a" and "b" positions would also not be fixed.), then
                    If LocalRow(1) > 0 And LocalRow(2) = 0 Then     ' If "b" is non-zero and "c' is zero, then switch "b" and "c".
                        LocalRow(2) = LocalRow(1)
                        LocalRow(1) = 0
                        If LocalRow(0) > 0 And LocalRow(1) = 0 Then  ' Repeat first If statement above.
                            LocalRow(1) = LocalRow(0)
                            LocalRow(0) = 0
                        End If
                    End If
                End If
            End If
            If FixedColumns(4) = 0 Then     ' Second MiniRow
                If LocalRow(3) > 0 And LocalRow(4) = 0 Then
                    LocalRow(4) = LocalRow(3)
                    LocalRow(3) = 0
                End If
                If FixedColumns(5) = 0 Then
                    If LocalRow(4) > 0 And LocalRow(5) = 0 Then
                        LocalRow(5) = LocalRow(4)
                        LocalRow(4) = 0
                        If LocalRow(3) > 0 And LocalRow(4) = 0 Then
                            LocalRow(4) = LocalRow(3)
                            LocalRow(3) = 0
                        End If
                    End If
                End If
            End If
            If FixedColumns(7) = 0 Then     ' Third MiniRow
                If LocalRow(6) > 0 And LocalRow(7) = 0 Then
                    LocalRow(7) = LocalRow(6)
                    LocalRow(6) = 0
                End If
            End If

            FirstNonZeroDigitPositionInRow = 9               ' Using 0-8 cell notation, 9 = empty row (all zeros).
            FirstNonZeroDigit = 10
            For i = 0 To 8                                   ' Identify first non-zero digit position in row (if any).
                If LocalRow(i) > 0 Then
                    FirstNonZeroDigitPositionInRow = i
                    FirstNonZeroDigit = LocalRow(i)
                    Exit For
                End If
            Next i
            CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 0
            CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = FirstNonZeroDigit
            If FirstNonZeroDigitPositionInRow < 9 Then    ' If non-empty row, idenitfy candidate first digits in row considering possible permutations (assuming row position 9 is fixed).
                ' Idenitfy candidate first digits in row considering possible permutations (assuming row position 9 is fixed).
                If FixedColumns(5) = 0 And MiniRowCount(0) > 0 And MiniRowCount(0) = MiniRowCount(1) Then ' Check for possible stack permutations.
                    Select Case MiniRowCount(0)
                        Case 1
                            CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(2)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(5)
                        Case 2
                            CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 3
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(1)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(2)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(2) = LocalRow(4)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(3) = LocalRow(5)
                        Case 3
                            CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 5
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(0)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(1)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(2) = LocalRow(2)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(3) = LocalRow(3)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(4) = LocalRow(4)
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(5) = LocalRow(5)
                    End Select
                Else
                    Select Case FirstNonZeroDigitPositionInRow          ' Check for possible column permutations. Note: The FixedColumns indicators are right justified within MiniRows.
                        Case 0
                            If LocalRow(1) > 0 And FixedColumns(1) = 0 Then                                                     ' if FixedColumns(1) = 0 then FixedColumns(0) = 0  - ?? fix this.
                                If LocalRow(2) > 0 And FixedColumns(2) = 0 Then
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 2
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(0)
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(1)
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(2) = LocalRow(2)
                                Else
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(0)
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(1)
                                End If
                            End If
                        Case 1
                            If LocalRow(2) > 0 And FixedColumns(2) = 0 Then
                                CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1
                                CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(1)
                                CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(2)
                            End If
                        Case 3
                            If LocalRow(4) > 0 And FixedColumns(4) = 0 Then
                                If LocalRow(5) > 0 AndAlso FixedColumns(5) = 0 Then
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 2
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(3)
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(4)
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(2) = LocalRow(5)
                                Else
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(3)
                                    CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(4)
                                End If
                            End If
                        Case 4
                            If LocalRow(5) > 0 And FixedColumns(5) = 0 Then
                                CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1
                                CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(4)
                                CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(5)
                            End If
                        Case 6
                            If LocalRow(7) > 0 And FixedColumns(7) = 0 Then
                                CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1
                                CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = LocalRow(6)
                                CandidateDigitsConsideringPossibleStackAndColumnPermutations(1) = LocalRow(7)
                            End If
                        Case Else
                            CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 0
                            CandidateDigitsConsideringPossibleStackAndColumnPermutations(0) = FirstNonZeroDigit
                    End Select
                End If

                FirstNonZeroDigitRelabeled = DigitsRelabelWrk(CandidateDigitsConsideringPossibleStackAndColumnPermutations(0))
                For i = 1 To CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx
                    If FirstNonZeroDigitRelabeled > DigitsRelabelWrk(CandidateDigitsConsideringPossibleStackAndColumnPermutations(i)) Then
                        FirstNonZeroDigitRelabeled = DigitsRelabelWrk(CandidateDigitsConsideringPossibleStackAndColumnPermutations(i))
                    End If
                Next i
            Else     ' empty row case
                FirstNonZeroDigitRelabeled = 10
            End If   'If FirstNonZeroDigitPositionInRow < 9 Then
        Else
            FirstNonZeroDigitPositionInRow = 9               ' Using 0-8 cell notation, 9 = empty row (all zeros).
            FirstNonZeroDigitRelabeled = 10
            For i = 0 To 8                                   ' Identify first non-zero digit position in row (if any).
                If LocalRow(i) > 0 Then
                    FirstNonZeroDigitPositionInRow = i
                    FirstNonZeroDigitRelabeled = DigitsRelabelWrk(LocalRow(i))
                    Exit For
                End If
            Next i
        End If
    End Sub

'   Shared Sub CalcRowPositionalWeight( Row As Integer,
'                                       StillJustifyingSw As Boolean,
'                                       Puzzle As Integer(),
'                                       FixedColumns As Integer(),
'                                 ByRef RowPositionalWeight As Integer)
'       Dim i, hold As Integer
'       Dim hold1, hold2, hold3 As Integer
'       '                                               0  1  2  3  4   5   6   7    8    9
'       Dim IntToBit1To9 As Integer() = New Integer(9) {0, 1, 2, 4, 8, 16, 32, 64, 128, 256}
'       Dim MiniRowCount() As Integer = New Integer(2) {}
'       Dim LocalRow As Integer() = New Integer(8) {}
'
'       Array.ConstrainedCopy(Puzzle, (Row - 1) * 9, LocalRow, 0, 9)
'
'       If StillJustifyingSw Then
'           For i = 0 To 8                                   ' Count non-zero digits in minirow.
'               If LocalRow(i) > 0 Then
'                   MiniRowCount(i \ 3) += 1
'               End If
'           Next i
'           ' "right justify" MiniRows in row.                    Note: The FixedColumns indicators are right justified within the row.
'           If FixedColumns(5) = 0 Then                         ' If minirow2 is not fixed, then minirow1 is also not fixed.
'               If MiniRowCount(0) > MiniRowCount(1) Then       ' If minirow1 count is greater than minirow2, switch minirow1 and minirow2.
'                   hold = MiniRowCount(0) : MiniRowCount(0) = MiniRowCount(1) : MiniRowCount(1) = hold
'                   hold1 = LocalRow(0)
'                   hold2 = LocalRow(1)
'                   hold3 = LocalRow(2)
'                   LocalRow(0) = LocalRow(3)
'                   LocalRow(1) = LocalRow(4)
'                   LocalRow(2) = LocalRow(5)
'                   LocalRow(3) = hold1
'                   LocalRow(4) = hold2
'                   LocalRow(5) = hold3
'               End If
'           End If
'
'           ' "right justify" columns in row within MiniRows. Note: The FixedColumn indicators are right justified within MiniRows.
'           ' For each MiniRow (0 to 2 notation). Note: the comments below use "abc" notation for digits in the MiniRow.
'           If FixedColumns(1) = 0 Then     ' First MiniRow            ' If the "b" position is not fixed (then the "a" position would also not be fixed.), then
'               If LocalRow(0) > 0 And LocalRow(1) = 0 Then         ' If "a" is non-zero and "b' is zero, then switch "a" and "b".
'                   LocalRow(1) = LocalRow(0)
'                   LocalRow(0) = 0
'               End If
'               If FixedColumns(2) = 0 Then                ' If "c" position is not fixed (then the "a" and "b" positions would also not be fixed.), then
'                   If LocalRow(1) > 0 And LocalRow(2) = 0 Then     ' If "b" is non-zero and "c' is zero, then switch "b" and "c".
'                       LocalRow(2) = LocalRow(1)
'                       LocalRow(1) = 0
'                       If LocalRow(0) > 0 And LocalRow(1) = 0 Then  ' Repeat first If statement above.
'                           LocalRow(1) = LocalRow(0)
'                           LocalRow(0) = 0
'                       End If
'                   End If
'               End If
'           End If
'           If FixedColumns(4) = 0 Then     ' Second MiniRow
'               If LocalRow(3) > 0 And LocalRow(4) = 0 Then
'                   LocalRow(4) = LocalRow(3)
'                   LocalRow(3) = 0
'               End If
'               If FixedColumns(5) = 0 Then
'                   If LocalRow(4) > 0 And LocalRow(5) = 0 Then
'                       LocalRow(5) = LocalRow(4)
'                       LocalRow(4) = 0
'                       If LocalRow(3) > 0 And LocalRow(4) = 0 Then
'                           LocalRow(4) = LocalRow(3)
'                           LocalRow(3) = 0
'                       End If
'                   End If
'               End If
'           End If
'           If FixedColumns(7) = 0 Then     ' Third MiniRow
'               If LocalRow(6) > 0 And LocalRow(7) = 0 Then
'                   LocalRow(7) = LocalRow(6)
'                   LocalRow(6) = 0
'               End If
'           End If
'       End If
'       RowPositionalWeight = 0
'       For i = 0 To 8
'           If LocalRow(i) > 0 Then
'               RowPositionalWeight = RowPositionalWeight Or IntToBit1To9(9 - i)
'           End If
'       Next i
'   End Sub
'
    Shared Sub SwitchStacks( Puzzle As Integer(), stackXin As Integer, stackYin As Integer)
        Dim temp As Integer: Dim stackX As Integer = stackXin*3: Dim stackY As Integer = stackYin*3
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 7: stackY += 7
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp: stackX += 1: stackY += 1
        temp = Puzzle(stackX): Puzzle(stackX) = Puzzle(stackY): Puzzle(stackY) = temp
    End Sub

    Shared Sub SwitchStacks01( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 0): Puzzle( 0) = Puzzle( 3): Puzzle( 3) = temp
        temp = Puzzle( 1): Puzzle( 1) = Puzzle( 4): Puzzle( 4) = temp
        temp = Puzzle( 2): Puzzle( 2) = Puzzle( 5): Puzzle( 5) = temp
        temp = Puzzle( 9): Puzzle( 9) = Puzzle(12): Puzzle(12) = temp
        temp = Puzzle(10): Puzzle(10) = Puzzle(13): Puzzle(13) = temp
        temp = Puzzle(11): Puzzle(11) = Puzzle(14): Puzzle(14) = temp
        temp = Puzzle(18): Puzzle(18) = Puzzle(21): Puzzle(21) = temp
        temp = Puzzle(19): Puzzle(19) = Puzzle(22): Puzzle(22) = temp
        temp = Puzzle(20): Puzzle(20) = Puzzle(23): Puzzle(23) = temp
        temp = Puzzle(27): Puzzle(27) = Puzzle(30): Puzzle(30) = temp
        temp = Puzzle(28): Puzzle(28) = Puzzle(31): Puzzle(31) = temp
        temp = Puzzle(29): Puzzle(29) = Puzzle(32): Puzzle(32) = temp
        temp = Puzzle(36): Puzzle(36) = Puzzle(39): Puzzle(39) = temp
        temp = Puzzle(37): Puzzle(37) = Puzzle(40): Puzzle(40) = temp
        temp = Puzzle(38): Puzzle(38) = Puzzle(41): Puzzle(41) = temp
        temp = Puzzle(45): Puzzle(45) = Puzzle(48): Puzzle(48) = temp
        temp = Puzzle(46): Puzzle(46) = Puzzle(49): Puzzle(49) = temp
        temp = Puzzle(47): Puzzle(47) = Puzzle(50): Puzzle(50) = temp
        temp = Puzzle(54): Puzzle(54) = Puzzle(57): Puzzle(57) = temp
        temp = Puzzle(55): Puzzle(55) = Puzzle(58): Puzzle(58) = temp
        temp = Puzzle(56): Puzzle(56) = Puzzle(59): Puzzle(59) = temp
        temp = Puzzle(63): Puzzle(63) = Puzzle(66): Puzzle(66) = temp
        temp = Puzzle(64): Puzzle(64) = Puzzle(67): Puzzle(67) = temp
        temp = Puzzle(65): Puzzle(65) = Puzzle(68): Puzzle(68) = temp
        temp = Puzzle(72): Puzzle(72) = Puzzle(75): Puzzle(75) = temp
        temp = Puzzle(73): Puzzle(73) = Puzzle(76): Puzzle(76) = temp
        temp = Puzzle(74): Puzzle(74) = Puzzle(77): Puzzle(77) = temp
    End Sub

    Shared Sub SwitchStacks02( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 0): Puzzle( 0) = Puzzle( 6): Puzzle( 6) = temp
        temp = Puzzle( 1): Puzzle( 1) = Puzzle( 7): Puzzle( 7) = temp
        temp = Puzzle( 2): Puzzle( 2) = Puzzle( 8): Puzzle( 8) = temp
        temp = Puzzle( 9): Puzzle( 9) = Puzzle(15): Puzzle(15) = temp
        temp = Puzzle(10): Puzzle(10) = Puzzle(16): Puzzle(16) = temp
        temp = Puzzle(11): Puzzle(11) = Puzzle(17): Puzzle(17) = temp
        temp = Puzzle(18): Puzzle(18) = Puzzle(24): Puzzle(24) = temp
        temp = Puzzle(19): Puzzle(19) = Puzzle(25): Puzzle(25) = temp
        temp = Puzzle(20): Puzzle(20) = Puzzle(26): Puzzle(26) = temp
        temp = Puzzle(27): Puzzle(27) = Puzzle(33): Puzzle(33) = temp
        temp = Puzzle(28): Puzzle(28) = Puzzle(34): Puzzle(34) = temp
        temp = Puzzle(29): Puzzle(29) = Puzzle(35): Puzzle(35) = temp
        temp = Puzzle(36): Puzzle(36) = Puzzle(42): Puzzle(42) = temp
        temp = Puzzle(37): Puzzle(37) = Puzzle(43): Puzzle(43) = temp
        temp = Puzzle(38): Puzzle(38) = Puzzle(44): Puzzle(44) = temp
        temp = Puzzle(45): Puzzle(45) = Puzzle(51): Puzzle(51) = temp
        temp = Puzzle(46): Puzzle(46) = Puzzle(52): Puzzle(52) = temp
        temp = Puzzle(47): Puzzle(47) = Puzzle(53): Puzzle(53) = temp
        temp = Puzzle(54): Puzzle(54) = Puzzle(60): Puzzle(60) = temp
        temp = Puzzle(55): Puzzle(55) = Puzzle(61): Puzzle(61) = temp
        temp = Puzzle(56): Puzzle(56) = Puzzle(62): Puzzle(62) = temp
        temp = Puzzle(63): Puzzle(63) = Puzzle(69): Puzzle(69) = temp
        temp = Puzzle(64): Puzzle(64) = Puzzle(70): Puzzle(70) = temp
        temp = Puzzle(65): Puzzle(65) = Puzzle(71): Puzzle(71) = temp
        temp = Puzzle(72): Puzzle(72) = Puzzle(78): Puzzle(78) = temp
        temp = Puzzle(73): Puzzle(73) = Puzzle(79): Puzzle(79) = temp
        temp = Puzzle(74): Puzzle(74) = Puzzle(80): Puzzle(80) = temp
    End Sub

    Shared Sub SwitchStacks12( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 3): Puzzle( 3) = Puzzle( 6): Puzzle( 6) = temp
        temp = Puzzle( 4): Puzzle( 4) = Puzzle( 7): Puzzle( 7) = temp
        temp = Puzzle( 5): Puzzle( 5) = Puzzle( 8): Puzzle( 8) = temp
        temp = Puzzle(12): Puzzle(12) = Puzzle(15): Puzzle(15) = temp
        temp = Puzzle(13): Puzzle(13) = Puzzle(16): Puzzle(16) = temp
        temp = Puzzle(14): Puzzle(14) = Puzzle(17): Puzzle(17) = temp
        temp = Puzzle(21): Puzzle(21) = Puzzle(24): Puzzle(24) = temp
        temp = Puzzle(22): Puzzle(22) = Puzzle(25): Puzzle(25) = temp
        temp = Puzzle(23): Puzzle(23) = Puzzle(26): Puzzle(26) = temp
        temp = Puzzle(30): Puzzle(30) = Puzzle(33): Puzzle(33) = temp
        temp = Puzzle(31): Puzzle(31) = Puzzle(34): Puzzle(34) = temp
        temp = Puzzle(32): Puzzle(32) = Puzzle(35): Puzzle(35) = temp
        temp = Puzzle(39): Puzzle(39) = Puzzle(42): Puzzle(42) = temp
        temp = Puzzle(40): Puzzle(40) = Puzzle(43): Puzzle(43) = temp
        temp = Puzzle(41): Puzzle(41) = Puzzle(44): Puzzle(44) = temp
        temp = Puzzle(48): Puzzle(48) = Puzzle(51): Puzzle(51) = temp
        temp = Puzzle(49): Puzzle(49) = Puzzle(52): Puzzle(52) = temp
        temp = Puzzle(50): Puzzle(50) = Puzzle(53): Puzzle(53) = temp
        temp = Puzzle(57): Puzzle(57) = Puzzle(60): Puzzle(60) = temp
        temp = Puzzle(58): Puzzle(58) = Puzzle(61): Puzzle(61) = temp
        temp = Puzzle(59): Puzzle(59) = Puzzle(62): Puzzle(62) = temp
        temp = Puzzle(66): Puzzle(66) = Puzzle(69): Puzzle(69) = temp
        temp = Puzzle(67): Puzzle(67) = Puzzle(70): Puzzle(70) = temp
        temp = Puzzle(68): Puzzle(68) = Puzzle(71): Puzzle(71) = temp
        temp = Puzzle(75): Puzzle(75) = Puzzle(78): Puzzle(78) = temp
        temp = Puzzle(76): Puzzle(76) = Puzzle(79): Puzzle(79) = temp
        temp = Puzzle(77): Puzzle(77) = Puzzle(80): Puzzle(80) = temp
    End Sub

    Shared Sub Switch3Stacks120( Puzzle As Integer())        ' 012 -> 120  rotate left
        Dim temp As Integer
        temp = Puzzle( 0): Puzzle( 0) = Puzzle( 3): Puzzle( 3) = Puzzle( 6): Puzzle( 6) = temp
        temp = Puzzle( 1): Puzzle( 1) = Puzzle( 4): Puzzle( 4) = Puzzle( 7): Puzzle( 7) = temp
        temp = Puzzle( 2): Puzzle( 2) = Puzzle( 5): Puzzle( 5) = Puzzle( 8): Puzzle( 8) = temp
        temp = Puzzle( 9): Puzzle( 9) = Puzzle(12): Puzzle(12) = Puzzle(15): Puzzle(15) = temp
        temp = Puzzle(10): Puzzle(10) = Puzzle(13): Puzzle(13) = Puzzle(16): Puzzle(16) = temp
        temp = Puzzle(11): Puzzle(11) = Puzzle(14): Puzzle(14) = Puzzle(17): Puzzle(17) = temp
        temp = Puzzle(18): Puzzle(18) = Puzzle(21): Puzzle(21) = Puzzle(24): Puzzle(24) = temp
        temp = Puzzle(19): Puzzle(19) = Puzzle(22): Puzzle(22) = Puzzle(25): Puzzle(25) = temp
        temp = Puzzle(20): Puzzle(20) = Puzzle(23): Puzzle(23) = Puzzle(26): Puzzle(26) = temp
        temp = Puzzle(27): Puzzle(27) = Puzzle(30): Puzzle(30) = Puzzle(33): Puzzle(33) = temp
        temp = Puzzle(28): Puzzle(28) = Puzzle(31): Puzzle(31) = Puzzle(34): Puzzle(34) = temp
        temp = Puzzle(29): Puzzle(29) = Puzzle(32): Puzzle(32) = Puzzle(35): Puzzle(35) = temp
        temp = Puzzle(36): Puzzle(36) = Puzzle(39): Puzzle(39) = Puzzle(42): Puzzle(42) = temp
        temp = Puzzle(37): Puzzle(37) = Puzzle(40): Puzzle(40) = Puzzle(43): Puzzle(43) = temp
        temp = Puzzle(38): Puzzle(38) = Puzzle(41): Puzzle(41) = Puzzle(44): Puzzle(44) = temp
        temp = Puzzle(45): Puzzle(45) = Puzzle(48): Puzzle(48) = Puzzle(51): Puzzle(51) = temp
        temp = Puzzle(46): Puzzle(46) = Puzzle(49): Puzzle(49) = Puzzle(52): Puzzle(52) = temp
        temp = Puzzle(47): Puzzle(47) = Puzzle(50): Puzzle(50) = Puzzle(53): Puzzle(53) = temp
        temp = Puzzle(54): Puzzle(54) = Puzzle(57): Puzzle(57) = Puzzle(60): Puzzle(60) = temp
        temp = Puzzle(55): Puzzle(55) = Puzzle(58): Puzzle(58) = Puzzle(61): Puzzle(61) = temp
        temp = Puzzle(56): Puzzle(56) = Puzzle(59): Puzzle(59) = Puzzle(62): Puzzle(62) = temp
        temp = Puzzle(63): Puzzle(63) = Puzzle(66): Puzzle(66) = Puzzle(69): Puzzle(69) = temp
        temp = Puzzle(64): Puzzle(64) = Puzzle(67): Puzzle(67) = Puzzle(70): Puzzle(70) = temp
        temp = Puzzle(65): Puzzle(65) = Puzzle(68): Puzzle(68) = Puzzle(71): Puzzle(71) = temp
        temp = Puzzle(72): Puzzle(72) = Puzzle(75): Puzzle(75) = Puzzle(78): Puzzle(78) = temp
        temp = Puzzle(73): Puzzle(73) = Puzzle(76): Puzzle(76) = Puzzle(79): Puzzle(79) = temp
        temp = Puzzle(74): Puzzle(74) = Puzzle(77): Puzzle(77) = Puzzle(80): Puzzle(80) = temp
    End Sub

    Shared Sub Switch3Stacks201( Puzzle As Integer())        ' 012 -> 210  rotate right
        Dim temp As Integer
        temp = Puzzle( 6): Puzzle( 6) = Puzzle( 3): Puzzle( 3) = Puzzle( 0): Puzzle( 0) = temp
        temp = Puzzle( 7): Puzzle( 7) = Puzzle( 4): Puzzle( 4) = Puzzle( 1): Puzzle( 1) = temp
        temp = Puzzle( 8): Puzzle( 8) = Puzzle( 5): Puzzle( 5) = Puzzle( 2): Puzzle( 2) = temp
        temp = Puzzle(15): Puzzle(15) = Puzzle(12): Puzzle(12) = Puzzle( 9): Puzzle( 9) = temp
        temp = Puzzle(16): Puzzle(16) = Puzzle(13): Puzzle(13) = Puzzle(10): Puzzle(10) = temp
        temp = Puzzle(17): Puzzle(17) = Puzzle(14): Puzzle(14) = Puzzle(11): Puzzle(11) = temp
        temp = Puzzle(24): Puzzle(24) = Puzzle(21): Puzzle(21) = Puzzle(18): Puzzle(18) = temp
        temp = Puzzle(25): Puzzle(25) = Puzzle(22): Puzzle(22) = Puzzle(19): Puzzle(19) = temp
        temp = Puzzle(26): Puzzle(26) = Puzzle(23): Puzzle(23) = Puzzle(20): Puzzle(20) = temp
        temp = Puzzle(33): Puzzle(33) = Puzzle(30): Puzzle(30) = Puzzle(27): Puzzle(27) = temp
        temp = Puzzle(34): Puzzle(34) = Puzzle(31): Puzzle(31) = Puzzle(28): Puzzle(28) = temp
        temp = Puzzle(35): Puzzle(35) = Puzzle(32): Puzzle(32) = Puzzle(29): Puzzle(29) = temp
        temp = Puzzle(42): Puzzle(42) = Puzzle(39): Puzzle(39) = Puzzle(36): Puzzle(36) = temp
        temp = Puzzle(43): Puzzle(43) = Puzzle(40): Puzzle(40) = Puzzle(37): Puzzle(37) = temp
        temp = Puzzle(44): Puzzle(44) = Puzzle(41): Puzzle(41) = Puzzle(38): Puzzle(38) = temp
        temp = Puzzle(51): Puzzle(51) = Puzzle(48): Puzzle(48) = Puzzle(45): Puzzle(45) = temp
        temp = Puzzle(52): Puzzle(52) = Puzzle(49): Puzzle(49) = Puzzle(46): Puzzle(46) = temp
        temp = Puzzle(53): Puzzle(53) = Puzzle(50): Puzzle(50) = Puzzle(47): Puzzle(47) = temp
        temp = Puzzle(60): Puzzle(60) = Puzzle(57): Puzzle(57) = Puzzle(54): Puzzle(54) = temp
        temp = Puzzle(61): Puzzle(61) = Puzzle(58): Puzzle(58) = Puzzle(55): Puzzle(55) = temp
        temp = Puzzle(62): Puzzle(62) = Puzzle(59): Puzzle(59) = Puzzle(56): Puzzle(56) = temp
        temp = Puzzle(69): Puzzle(69) = Puzzle(66): Puzzle(66) = Puzzle(63): Puzzle(63) = temp
        temp = Puzzle(70): Puzzle(70) = Puzzle(67): Puzzle(67) = Puzzle(64): Puzzle(64) = temp
        temp = Puzzle(71): Puzzle(71) = Puzzle(68): Puzzle(68) = Puzzle(65): Puzzle(65) = temp
        temp = Puzzle(78): Puzzle(78) = Puzzle(75): Puzzle(75) = Puzzle(72): Puzzle(72) = temp
        temp = Puzzle(79): Puzzle(79) = Puzzle(76): Puzzle(76) = Puzzle(73): Puzzle(73) = temp
        temp = Puzzle(80): Puzzle(80) = Puzzle(77): Puzzle(77) = Puzzle(74): Puzzle(74) = temp
    End Sub

    Shared Sub SwitchColumns( Puzzle As Integer(), columnX As Integer, columnY As Integer)
        Dim temp As Integer
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp: columnX += 9: columnY += 9
        temp = Puzzle(columnX): Puzzle(columnX) = Puzzle(columnY): Puzzle(columnY) = temp
    End Sub

    Shared Sub SwitchColumns01( Puzzle As Integer())
        Dim temp As Integer

        temp = Puzzle( 0): Puzzle( 0) = Puzzle( 1): Puzzle( 1) = temp
        temp = Puzzle( 9): Puzzle( 9) = Puzzle(10): Puzzle(10) = temp
        temp = Puzzle(18): Puzzle(18) = Puzzle(19): Puzzle(19) = temp
        temp = Puzzle(27): Puzzle(27) = Puzzle(28): Puzzle(28) = temp
        temp = Puzzle(36): Puzzle(36) = Puzzle(37): Puzzle(37) = temp
        temp = Puzzle(45): Puzzle(45) = Puzzle(46): Puzzle(46) = temp
        temp = Puzzle(54): Puzzle(54) = Puzzle(55): Puzzle(55) = temp
        temp = Puzzle(63): Puzzle(63) = Puzzle(64): Puzzle(64) = temp
        temp = Puzzle(72): Puzzle(72) = Puzzle(73): Puzzle(73) = temp
    End Sub

    Shared Sub SwitchColumns02( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 0): Puzzle( 0) = Puzzle( 2): Puzzle( 2) = temp
        temp = Puzzle( 9): Puzzle( 9) = Puzzle(11): Puzzle(11) = temp
        temp = Puzzle(18): Puzzle(18) = Puzzle(20): Puzzle(20) = temp
        temp = Puzzle(27): Puzzle(27) = Puzzle(29): Puzzle(29) = temp
        temp = Puzzle(36): Puzzle(36) = Puzzle(38): Puzzle(38) = temp
        temp = Puzzle(45): Puzzle(45) = Puzzle(47): Puzzle(47) = temp
        temp = Puzzle(54): Puzzle(54) = Puzzle(56): Puzzle(56) = temp
        temp = Puzzle(63): Puzzle(63) = Puzzle(65): Puzzle(65) = temp
        temp = Puzzle(72): Puzzle(72) = Puzzle(74): Puzzle(74) = temp
    End Sub

    Shared Sub SwitchColumns12( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 1): Puzzle( 1) = Puzzle( 2): Puzzle( 2) = temp
        temp = Puzzle(10): Puzzle(10) = Puzzle(11): Puzzle(11) = temp
        temp = Puzzle(19): Puzzle(19) = Puzzle(20): Puzzle(20) = temp
        temp = Puzzle(28): Puzzle(28) = Puzzle(29): Puzzle(29) = temp
        temp = Puzzle(37): Puzzle(37) = Puzzle(38): Puzzle(38) = temp
        temp = Puzzle(46): Puzzle(46) = Puzzle(47): Puzzle(47) = temp
        temp = Puzzle(55): Puzzle(55) = Puzzle(56): Puzzle(56) = temp
        temp = Puzzle(64): Puzzle(64) = Puzzle(65): Puzzle(65) = temp
        temp = Puzzle(73): Puzzle(73) = Puzzle(74): Puzzle(74) = temp
    End Sub

    Shared Sub SwitchColumns34( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 3): Puzzle( 3) = Puzzle( 4): Puzzle( 4) = temp
        temp = Puzzle(12): Puzzle(12) = Puzzle(13): Puzzle(13) = temp
        temp = Puzzle(21): Puzzle(21) = Puzzle(22): Puzzle(22) = temp
        temp = Puzzle(30): Puzzle(30) = Puzzle(31): Puzzle(31) = temp
        temp = Puzzle(39): Puzzle(39) = Puzzle(40): Puzzle(40) = temp
        temp = Puzzle(48): Puzzle(48) = Puzzle(49): Puzzle(49) = temp
        temp = Puzzle(57): Puzzle(57) = Puzzle(58): Puzzle(58) = temp
        temp = Puzzle(66): Puzzle(66) = Puzzle(67): Puzzle(67) = temp
        temp = Puzzle(75): Puzzle(75) = Puzzle(76): Puzzle(76) = temp
    End Sub

    Shared Sub SwitchColumns35( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 3): Puzzle( 3) = Puzzle( 5): Puzzle( 5) = temp
        temp = Puzzle(12): Puzzle(12) = Puzzle(14): Puzzle(14) = temp
        temp = Puzzle(21): Puzzle(21) = Puzzle(23): Puzzle(23) = temp
        temp = Puzzle(30): Puzzle(30) = Puzzle(32): Puzzle(32) = temp
        temp = Puzzle(39): Puzzle(39) = Puzzle(41): Puzzle(41) = temp
        temp = Puzzle(48): Puzzle(48) = Puzzle(50): Puzzle(50) = temp
        temp = Puzzle(57): Puzzle(57) = Puzzle(59): Puzzle(59) = temp
        temp = Puzzle(66): Puzzle(66) = Puzzle(68): Puzzle(68) = temp
        temp = Puzzle(75): Puzzle(75) = Puzzle(77): Puzzle(77) = temp
    End Sub

    Shared Sub SwitchColumns45( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 4): Puzzle( 4) = Puzzle( 5): Puzzle( 5) = temp
        temp = Puzzle(13): Puzzle(13) = Puzzle(14): Puzzle(14) = temp
        temp = Puzzle(22): Puzzle(22) = Puzzle(23): Puzzle(23) = temp
        temp = Puzzle(31): Puzzle(31) = Puzzle(32): Puzzle(32) = temp
        temp = Puzzle(40): Puzzle(40) = Puzzle(41): Puzzle(41) = temp
        temp = Puzzle(49): Puzzle(49) = Puzzle(50): Puzzle(50) = temp
        temp = Puzzle(58): Puzzle(58) = Puzzle(59): Puzzle(59) = temp
        temp = Puzzle(67): Puzzle(67) = Puzzle(68): Puzzle(68) = temp
        temp = Puzzle(76): Puzzle(76) = Puzzle(77): Puzzle(77) = temp
    End Sub

    Shared Sub SwitchColumns67( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 6): Puzzle( 6) = Puzzle( 7): Puzzle( 7) = temp
        temp = Puzzle(15): Puzzle(15) = Puzzle(16): Puzzle(16) = temp
        temp = Puzzle(24): Puzzle(24) = Puzzle(25): Puzzle(25) = temp
        temp = Puzzle(33): Puzzle(33) = Puzzle(34): Puzzle(34) = temp
        temp = Puzzle(42): Puzzle(42) = Puzzle(43): Puzzle(43) = temp
        temp = Puzzle(51): Puzzle(51) = Puzzle(52): Puzzle(52) = temp
        temp = Puzzle(60): Puzzle(60) = Puzzle(61): Puzzle(61) = temp
        temp = Puzzle(69): Puzzle(69) = Puzzle(70): Puzzle(70) = temp
        temp = Puzzle(78): Puzzle(78) = Puzzle(79): Puzzle(79) = temp
    End Sub

    Shared Sub SwitchColumns68( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 6): Puzzle( 6) = Puzzle( 8): Puzzle( 8) = temp
        temp = Puzzle(15): Puzzle(15) = Puzzle(17): Puzzle(17) = temp
        temp = Puzzle(24): Puzzle(24) = Puzzle(26): Puzzle(26) = temp
        temp = Puzzle(33): Puzzle(33) = Puzzle(35): Puzzle(35) = temp
        temp = Puzzle(42): Puzzle(42) = Puzzle(44): Puzzle(44) = temp
        temp = Puzzle(51): Puzzle(51) = Puzzle(53): Puzzle(53) = temp
        temp = Puzzle(60): Puzzle(60) = Puzzle(62): Puzzle(62) = temp
        temp = Puzzle(69): Puzzle(69) = Puzzle(71): Puzzle(71) = temp
        temp = Puzzle(78): Puzzle(78) = Puzzle(80): Puzzle(80) = temp
    End Sub

    Shared Sub SwitchColumns78( Puzzle As Integer())
        Dim temp As Integer
        temp = Puzzle( 7): Puzzle( 7) = Puzzle( 8): Puzzle( 8) = temp
        temp = Puzzle(16): Puzzle(16) = Puzzle(17): Puzzle(17) = temp
        temp = Puzzle(25): Puzzle(25) = Puzzle(26): Puzzle(26) = temp
        temp = Puzzle(34): Puzzle(34) = Puzzle(35): Puzzle(35) = temp
        temp = Puzzle(43): Puzzle(43) = Puzzle(44): Puzzle(44) = temp
        temp = Puzzle(52): Puzzle(52) = Puzzle(53): Puzzle(53) = temp
        temp = Puzzle(61): Puzzle(61) = Puzzle(62): Puzzle(62) = temp
        temp = Puzzle(70): Puzzle(70) = Puzzle(71): Puzzle(71) = temp
        temp = Puzzle(79): Puzzle(79) = Puzzle(80): Puzzle(80) = temp
    End Sub

    Shared Sub Switch3Columns120( Puzzle As Integer())        ' 012 -> 120  rotate left
        Dim temp As Integer
        temp = Puzzle( 0): Puzzle( 0) = Puzzle( 1): Puzzle( 1) = Puzzle( 2): Puzzle( 2) = temp
        temp = Puzzle( 9): Puzzle( 9) = Puzzle(10): Puzzle(10) = Puzzle(11): Puzzle(11) = temp
        temp = Puzzle(18): Puzzle(18) = Puzzle(19): Puzzle(19) = Puzzle(20): Puzzle(20) = temp
        temp = Puzzle(27): Puzzle(27) = Puzzle(28): Puzzle(28) = Puzzle(29): Puzzle(29) = temp
        temp = Puzzle(36): Puzzle(36) = Puzzle(37): Puzzle(37) = Puzzle(38): Puzzle(38) = temp
        temp = Puzzle(45): Puzzle(45) = Puzzle(46): Puzzle(46) = Puzzle(47): Puzzle(47) = temp
        temp = Puzzle(54): Puzzle(54) = Puzzle(55): Puzzle(55) = Puzzle(56): Puzzle(56) = temp
        temp = Puzzle(63): Puzzle(63) = Puzzle(64): Puzzle(64) = Puzzle(65): Puzzle(65) = temp
        temp = Puzzle(72): Puzzle(72) = Puzzle(73): Puzzle(73) = Puzzle(74): Puzzle(74) = temp
    End Sub

    Shared Sub Switch3Columns201( Puzzle As Integer())        ' 012 -> 201  rotate right
        Dim temp As Integer
        temp = Puzzle( 2): Puzzle( 2) = Puzzle( 1): Puzzle( 1) = Puzzle( 0): Puzzle( 0) = temp
        temp = Puzzle(11): Puzzle(11) = Puzzle(10): Puzzle(10) = Puzzle( 9): Puzzle( 9) = temp
        temp = Puzzle(20): Puzzle(20) = Puzzle(19): Puzzle(19) = Puzzle(18): Puzzle(18) = temp
        temp = Puzzle(29): Puzzle(29) = Puzzle(28): Puzzle(28) = Puzzle(27): Puzzle(27) = temp
        temp = Puzzle(38): Puzzle(38) = Puzzle(37): Puzzle(37) = Puzzle(36): Puzzle(36) = temp
        temp = Puzzle(47): Puzzle(47) = Puzzle(46): Puzzle(46) = Puzzle(45): Puzzle(45) = temp
        temp = Puzzle(56): Puzzle(56) = Puzzle(55): Puzzle(55) = Puzzle(54): Puzzle(54) = temp
        temp = Puzzle(65): Puzzle(65) = Puzzle(64): Puzzle(64) = Puzzle(63): Puzzle(63) = temp
        temp = Puzzle(74): Puzzle(74) = Puzzle(73): Puzzle(73) = Puzzle(72): Puzzle(72) = temp
    End Sub

    Shared Sub Switch3Columns786( Puzzle As Integer())        ' 678 -> 786  rotate left
        Dim temp As Integer
        temp = Puzzle( 6): Puzzle( 6) = Puzzle( 7): Puzzle( 7) = Puzzle( 8): Puzzle( 8) = temp
        temp = Puzzle(15): Puzzle(15) = Puzzle(16): Puzzle(16) = Puzzle(17): Puzzle(17) = temp
        temp = Puzzle(24): Puzzle(24) = Puzzle(25): Puzzle(25) = Puzzle(26): Puzzle(26) = temp
        temp = Puzzle(33): Puzzle(33) = Puzzle(34): Puzzle(34) = Puzzle(35): Puzzle(35) = temp
        temp = Puzzle(42): Puzzle(42) = Puzzle(43): Puzzle(43) = Puzzle(44): Puzzle(44) = temp
        temp = Puzzle(51): Puzzle(51) = Puzzle(52): Puzzle(52) = Puzzle(53): Puzzle(53) = temp
        temp = Puzzle(60): Puzzle(60) = Puzzle(61): Puzzle(61) = Puzzle(62): Puzzle(62) = temp
        temp = Puzzle(69): Puzzle(69) = Puzzle(70): Puzzle(70) = Puzzle(71): Puzzle(71) = temp
        temp = Puzzle(78): Puzzle(78) = Puzzle(79): Puzzle(79) = Puzzle(80): Puzzle(80) = temp
    End Sub

    Shared Sub Switch3Columns867( Puzzle As Integer())        ' 678 -> 867  rotate right
        Dim temp As Integer
        temp = Puzzle( 8): Puzzle( 8) = Puzzle( 7): Puzzle( 7) = Puzzle( 6): Puzzle( 6) = temp
        temp = Puzzle(17): Puzzle(17) = Puzzle(16): Puzzle(16) = Puzzle(15): Puzzle(15) = temp
        temp = Puzzle(26): Puzzle(26) = Puzzle(25): Puzzle(25) = Puzzle(24): Puzzle(24) = temp
        temp = Puzzle(35): Puzzle(35) = Puzzle(34): Puzzle(34) = Puzzle(33): Puzzle(33) = temp
        temp = Puzzle(44): Puzzle(44) = Puzzle(43): Puzzle(43) = Puzzle(42): Puzzle(42) = temp
        temp = Puzzle(53): Puzzle(53) = Puzzle(52): Puzzle(52) = Puzzle(51): Puzzle(51) = temp
        temp = Puzzle(62): Puzzle(62) = Puzzle(61): Puzzle(61) = Puzzle(60): Puzzle(60) = temp
        temp = Puzzle(71): Puzzle(71) = Puzzle(70): Puzzle(70) = Puzzle(69): Puzzle(69) = temp
        temp = Puzzle(80): Puzzle(80) = Puzzle(79): Puzzle(79) = Puzzle(78): Puzzle(78) = temp
    End Sub

End Class

