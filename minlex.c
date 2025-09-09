/* minlex.c - mostly converted from MinLexDriver.vb
*/
const char *AUTHOR  = "1to9only";
const char *VERSION = "2025.09.09";
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <fcntl.h>
#include <io.h>
/* 9 x 9 */
#define SUDOKU_SIZE  9
#define SUDOKU_MIN  '0'
#define SUDOKU_ONE  '1'
#define SUDOKU_MAX  '9'
#define SUDOKU_DOT  '.'
#define GRID_SIZE    81
#define LINE_SIZE    83

#define BATCH_SIZE   1296
int  total = 0, count = 0;                // counts
char sudokus[ BATCH_SIZE][ LINE_SIZE];    // sudokus
char canons[ BATCH_SIZE][ LINE_SIZE];     // canons

int MinLex9X9SR1(char InputGridBufferChr[], int InputBufferSize, char MinLexBufferChr[], int ResultsBufferOffset);

#define NUM_THREADS  6
#define NUM_MINLEX   216

void batch_minlex( void)
{
   #pragma omp parallel for
   for (int index=0; index<NUM_THREADS; index++ )
   {
      int num = count - ( index * NUM_MINLEX);
      if ( num > NUM_MINLEX ) { num = NUM_MINLEX; }
      if ( num < 0 ) { num = 0; }
      if ( num != 0 )
      {
         MinLex9X9SR1( (char*)sudokus[ index*NUM_MINLEX], num*LINE_SIZE, (char*)canons[ index*NUM_MINLEX], 0);
      }
   }
}

int main( int argc, char *argv[])
{
   char buffer[ 1024] = { 0, };
   char *sudoku = buffer;

   if ( argc == 1 )
   {
      printf( "minlex.exe [-v | puzzle | file_containing_sudokus]\ne.g.\n");
      printf( "minlex.exe ......789......1.....2.3...2......65.9..71.............7..9.8..............5...4.\n");
      printf( "minlex.exe 17puz49158.txt > output.txt\n");
      exit( 0);
   }
   for (register int arg=1; arg<argc; arg++ )
   {
      if ( ( argv[arg][0]=='-' || argv[arg][0]=='/' )
        && ( argv[arg][1]=='v' || argv[arg][1]=='V' )
          && argv[arg][2]==0 )
      {
#ifndef _OPENMP
         printf( "VERSION: %s, single_thread.\n", VERSION);
#endif
#ifdef _OPENMP
         printf( "VERSION: %s, multi_threads.\n", VERSION);
#endif
         exit( 0);
      }
      if ( strlen( argv[arg]) == GRID_SIZE )    // is sudoku line
      {
         sudoku = argv[arg];
         count = 0;
         for (register int i=0; i<GRID_SIZE; i++ )
         {
            if ( sudoku[ i] >= SUDOKU_ONE && sudoku[ i] <= SUDOKU_MAX )
            {
               sudokus[ count][ i] = sudoku[ i];
            }
            else
            {
               sudokus[ count][ i] = SUDOKU_MIN;
            }
         }
         count += 1;
         batch_minlex();
         for ( register int i=0; i<count; i++ )
         {
            printf( "%s\n", canons[ i]);
         }
      }
      else                                      // is sudoku file
      {
         struct _stat st = { 0, };
         if ( _stat( argv[arg], &st) == 0 )
         {
            FILE *f = fopen( argv[arg], "r");
            if ( f != (FILE*)0 )
            {
               total = 0;
               count = 0;
               while( fgets( sudoku, 1020, f) )          // read one line
               {
                  sudoku[ strcspn( sudoku, "\r\n")] = 0; // remove CR, LF, CRLF, LFCR, ...
                  char ch = sudoku[0]; // examine first character of line
                  // ignore lines starting with these chars
                  if ( !( ch == 0 || ch==' ' || ch=='_' || ch=='#' || (ch>='A' && ch<='Z') || (ch>='a' && ch<='z')) )
                  if ( strlen( sudoku) >= GRID_SIZE )
                  {
                     for (register int i=0; i<GRID_SIZE; i++ )
                     {
                        if ( sudoku[ i] >= SUDOKU_ONE && sudoku[ i] <= SUDOKU_MAX )
                        {
                           sudokus[ count][ i] = sudoku[ i];
                        }
                        else
                        {
                           sudokus[ count][ i] = SUDOKU_MIN;
                        }
                     }
                     count += 1;

                     if ( count == BATCH_SIZE )
                     {
                        total += count;
                        batch_minlex();
                        for ( register int i=0; i<count; i++ )
                        {
                           printf( "%s\n", canons[ i]);
                        }
                        count = 0;
                     }
                  }
               }
               fclose( f);

               if ( count != 0 )
               {
                  total += count;
                  batch_minlex();
                  for ( register int i=0; i<count; i++ )
                  {
                     printf( "%s\n", canons[ i]);
                  }
                  count = 0;
               }

            }
         }
         else
         {
            printf( "file '%s' not found!\n", argv[arg]);
            exit( 1);
         }
      }
   }

   return 0;
}

void RightJustifyRow(int Row, int Puzzle[], bool &StillJusfifyingSw, int FixedColumns[], int &StackPermutationCode, int &ColumnPermutationCode);
void FindFirstNonZeroDigitInRow(int Row, bool StillJustifyingSw, int Puzzle[], int FixedColumns[], int DigitsRelabelWrk[], int &FirstNonZeroDigitPositionInRow, int &FirstNonZeroDigitRelabeled);

void SwitchStacksXY( int Puzzle[], int stackXin, int stackYin);
void SwitchStacks01( int Puzzle[]);
void SwitchStacks02( int Puzzle[]);
void SwitchStacks12( int Puzzle[]);
void Switch3Stacks120( int Puzzle[]);  // 012 -> 120  rotate left
void Switch3Stacks201( int Puzzle[]);  // 012 -> 210  rotate right

void SwitchColumnsXY( int Puzzle[], int columnX, int columnY);
void SwitchColumns01( int Puzzle[]);
void SwitchColumns02( int Puzzle[]);
void SwitchColumns12( int Puzzle[]);
void SwitchColumns34( int Puzzle[]);
void SwitchColumns35( int Puzzle[]);
void SwitchColumns45( int Puzzle[]);
void SwitchColumns67( int Puzzle[]);
void SwitchColumns68( int Puzzle[]);
void SwitchColumns78( int Puzzle[]);
void Switch3Columns120( int Puzzle[]); // 012 -> 120  rotate left
void Switch3Columns201( int Puzzle[]); // 012 -> 201  rotate right
void Switch3Columns786( int Puzzle[]); // 678 -> 786  rotate left
void Switch3Columns867( int Puzzle[]); // 678 -> 867  rotate right

static inline void ArrayClear( int array[], int index, int length)
{
   for (int i=0; i<length; i++ ) { array[ index+i] = 0; }
}
static inline void ArrayConstrainedCopy( int srcArray[], int srcIndex, int tgtArray[], int tgtIndex, int length)
{
   for (int i=0; i<length; i++ ) { tgtArray[ tgtIndex+i] = srcArray[ srcIndex+i]; }
}

int MinLex9X9SR1(char InputGridBufferChr[], int InputBufferSize, char MinLexBufferChr[], int ResultsBufferOffset)
{
   int a = 0;
   int b = 0;
   int i = 0;
   int j = 0;
   int k = 0;
   int m = 0;
   int n = 0;
   int z = 0;
   int zstart = 0;
   int hold = 0;

   int InputGrid[162];
   int TestBand1[81];
   int TestBand2[81];
   int TestBand3[81];
   int TestGridRelabeled[81];
   int HoldPuzzle[81];
   int HoldGrid[81];
   int HoldRow[9];
   int iForHit[9];

   int FirstNonZeroDigitPositionInRow[10]; //NOTE: Zero & 1 elements not used.
   int FirstNonZeroDigitRelabeled[10];     //NOTE: Zero, 1, 2, 3 elements not used.
   int CandidateFirstRelabeledDigit = 0;

   int ColumnPermutationTrackerIx = 0;
   int CandidateColumnPermutationTrackerStartIx = 0;
  bool FirstColumnPermutationTrackerIsIdentitySw = false;
  bool ColumnPermutationTrackerIsIdentitySw[1000]; // 1000 - 1 ' Identifies those Trackers that indicate no change (The "identity" permutation - "012345678")
   int ColumnPermutationTracker[9000];             // 1000 X 9 - 1
  bool HoldDigitAlreadyHitSw[360];                 // 36 X 10 - 1
   int HoldDigitsRelabelWrk[10000];                // 1000 X 10 - 1
   int HoldRelabelLastDigit[36];                   // 36 - 1
   int HoldBand1CandidateJustifiedPuzzles[2916];   // 36 X 81 - 1

   int AugmentedMiniRowCount1 = 0;
   int AugmentedMiniRowCount2 = 0;
   int CalcMiniRowCountCodeMinimum = 0;
   int CalcMiniRowCountCode[18];
  bool CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = false;
  bool CheckThisPassSw = false;
   int ColumnPermutationCode = 0;
  bool DigitAlreadyHitSw[10];     // Note: zero element not used.
   int DigitsInMiniRowBit[18][3];
   int DigitsInRowBit[18];
   int FirstNonZeroRowCandidateIx = 0;
   int FoundColumnPermutationTrackerCountMax = 0;
   int JustifiedMiniRowCount[18][3];
   int LocalBand1[27];
   int LocalRows2and3[19];
   int MiniRowRepeatDigits = 0;
  bool MinLexCandidateSw = false;
  bool MinLexRowHitSw = false;
   int MinLexRows[27];
   int OriginalMiniRowCount[18][3];
   int RelabelLastDigit = 0;
  bool ResetMinLexRowSw = false;
   int StackPermutationTrackerCode = 0;
   int StackPermutationTrackerLocal[3];
   int stackx = 0;
   int stacky = 0;
   int StartEqualCheck = 0;
  bool StillJustifyingSw = false;
   int StoppedJustifyingRow = 0;
   int switchx = 0;
   int switchy = 0;
   int TestRowsRelabeled[27];
   int ZeroRowsInBandsCount[6];

   int Band1CandidateIx = 0;
   int Band1CandidateRow1[36];
   int Band1CandidateRow2[36];
   int Band1CandidateRow3[36];
   int Band1MiniRowColumnPermutationCode[3];

   int FixedColumns[9]; // Indicates if a column containes a non-zero digit above the current row. Elements 0 to 8 correspond to columns 1 to 9.
  bool FixedColumnsSavedAsOfRow4Sw = false;
  bool FixedColumnsSavedAsOfRow6Sw = false;
  bool FixedColumnsSavedAsOfRow7Sw = false;

   int Step2Row1CandidateIx = 0;
   int Step2Row1Candidate[36];
   int Step2Row2Candidate[36];
   int Step2ColumnPermutationTrackerStartIx[36];
   int Step2ColumnPermutationTrackerCount[36];
   int ColumnPermutationTrackerCount = 0;

   int TwoRowCandidateIx = 0;
   int TwoRowCandidateRow1[36];
   int TwoRowCandidateRow2[36];
   int TwoRowCandidateRow3[36];
   int TwoRowCandidateMiniRowCode[36];
   int TwoRowCandidateMiniRow1Count = 0;
   int TwoRowCandidateMiniRow2Count = 0;
   int TwoRowCandidateMiniRowCodeMinimum = 0;
   int TwoRowCandidateMiniRow1CodeMinimum = 0;

   int RowWithCalcMiniRowCountCodeMinimumIx = 0;
   int RowWithCalcMiniRowCountCodeMinimum[18];
   int RowRepeatDigits = 0;
   int Row2Candidate = 0;
   int Row2CalcMiniRowCountCodeMinimum = 0;
   int Row1StackPermutationCode = 0;
   int Row1Candidate = 0;
   int Row1Candidate17 = 0;
   int Row2Candidate17 = 0;
   int Row3Candidate = 0;
   int row1start = 0;
   int row2start = 0;
   int row3start = 0;
   int Row1MiniRowCount[3];
   int Row2MiniRowCount[3];
   int Row2Or3StackPermutationCode = 0;
   int Row3MiniRowCount[3];
   int Row2TestFirstNonZeroDigitPositionInRow = 0;
   int Row2StackPermutationCode = 0;
   int Row3Candidate17 = 0;
   int Row3StackPermutationCode = 0;
   int Row3TestFirstNonZeroDigitPositionInRow = 0;
  bool Row3DigitAlreadyHitSw[10]; // Note: zero element not used.
   int Row3RelabelLastDigit = 0;
   int Row4TestFirstNonZeroDigitPositionInRow = 0;
   int Row4TestPositionalCandidateRow[10];
   int Row4TestCandidateRowIx = 0;
   int Row4TestCandidateRow[9];
   int Row3FixedColumns[9];
   int Row4StackPermutationCode = 0;
   int Row4ColumnPermutationCode = 0;
   int Row4FixedColumns[9];
  bool Row4DigitAlreadyHitSw[10]; // Note: zero element not used.
   int Row4RelabelLastDigit = 0;
   int Row5TestFirstNonZeroDigitPositionInRow = 0;
   int Row5TestCandidateRowIx = 0;
   int Row5TestCandidateRow[2];
   int Row5StackPermutationCode = 0;
   int Row5ColumnPermutationCode = 0;
   int Row5FixedColumns[9];
  bool Row5DigitAlreadyHitSw[10]; // Note: zero element not used.
   int Row5RelabelLastDigit = 0;
   int Row6StackPermutationCode = 0;
   int Row6ColumnPermutationCode = 0;
   int Row6FixedColumns[9];
  bool Row6DigitAlreadyHitSw[10]; // Note: zero element not used.
   int Row6RelabelLastDigit = 0;
   int Row7TestFirstNonZeroDigitPositionInRow = 0;
   int Row7TestPositionalCandidateRow[10];
   int Row7TestCandidateRowIx = 0;
   int Row7TestCandidateRow[3];
   int Row7StackPermutationCode = 0;
   int Row7ColumnPermutationCode = 0;
   int Row7FixedColumns[9];
  bool Row7DigitAlreadyHitSw[10]; // Note: zero element not used.
   int Row7RelabelLastDigit = 0;
   int Row8TestFirstNonZeroDigitPositionInRow = 0;
   int Row8TestCandidateRowIx = 0;
   int Row8TestCandidateRow[2];
   int Row8StackPermutationCode = 0;
   int Row8ColumnPermutationCode = 0;
   int Row8FixedColumns[9];
  bool Row8DigitAlreadyHitSw[10]; // Note: zero element not used.
   int Row8RelabelLastDigit = 0;
   int Row9StackPermutationCode = 0;
   int Row9ColumnPermutationCode = 0;
  bool Row9DigitAlreadyHitSw[10]; // Note: zero element not used.
   int Row9RelabelLastDigit = 0;

   static char CharPeriodTO9[10] = { '.', '1', '2', '3', '4', '5', '6', '7', '8', '9'};
   static int IntToBit1To9[10] = { 0, 1, 2, 4, 8, 16, 32, 64, 128, 256};

   static int ColumnTrackerInit[9] = { 0, 1, 2, 3, 4, 5, 6, 7, 8};
   int LocalColumnPermutationTracker[9];

   // For DigitsRelabelWrk, the zero element relabels 0 to 0, the 10 element is used to assign the first digit in an empty row to position "10".
   static int DigitsRelabelWrkInit[11] = { 0, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10};
   int DigitsRelabelWrk[11];

   int Row3DigitsRelabelWrk[11];
   int Row4DigitsRelabelWrk[11];
   int Row5DigitsRelabelWrk[11];
   int Row6DigitsRelabelWrk[11];
   int Row7DigitsRelabelWrk[11];
   int Row8DigitsRelabelWrk[11];
   int Row9DigitsRelabelWrk[11];

   static int StackPermutations[4] = { 0, 1, 1, 5};

   static int MinLexFirstNonZeroDigitPositionInRowReset[10] = { 0, 0, -1, -1, -1, -1, -1, -1, -1, -1}; //NOTE: Zero, 1, 2, 3 elements not used.
   int MinLexFirstNonZeroDigitPositionInRow[10] = { 0, 0, -1, -1, -1, -1, -1, -1, -1, -1};

   static int MinLexFullGridLocalReset[81] = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 4, 5, 7, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9};
   static int MinLexGridLocalReset[81] = { 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10};
   int MinLexGridLocal[81];

   //       Row 1 Candidate Rows =   1  2  3  4  5  6  7  8  9  10  11 12  13  14  15  16  17  18
   static int Row2CandidateA[18] = { 1, 2, 0, 4, 5, 3, 7, 8, 6, 10, 11, 9, 13, 14, 12, 16, 17, 15}; // Note: in 0 to 17 notation.
   static int Row2CandidateB[18] = { 2, 0, 1, 5, 3, 4, 8, 6, 7, 11, 9, 10, 14, 12, 13, 17, 15, 16};

   //                                                       0  1  2  3   4  5  6  7   8   9 10 11 12  13  14  15  16  17  18  19  20 21 22 23  24  25 26 27  28  29  30 31  32  33  34  35  36  37  38  39  40  41 42 43  44  45  46 47  48  49  50  51  52  53  54  55  56  57  58  59  60  61  62 63
   //                                                          |  |  |      |  |  |          |  |              |                      |  |  |          |  |              |                                          |  |              |                                                              |
   static int FirstNonZeroPositionInFirstNonZeroRow[64] = {-1, 8, 7, 6, -1, 5, 5, 5, -1, -1, 4, 4, -1, -1, -1, 3, -1, -1, -1, -1, -1, 2, 2, 2, -1, -1, 2, 2, -1, -1, -1, 2, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 1, 1, -1, -1, -1, 1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 0};
   int FirstNonZeroPositionInRow1 = 0;
   int FirstNonZeroPositionInRow2 = 0;
   //                                                     0  1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40 41 42 43 44 45 46 47 48 49 50 51 52 53 54 55 56 57 58 59 60 61 62 63
   //                                                                    |              |              |                 |  |  |        |              |                                |  |           |                                               |
   static int FirstNonZeroRowStackPermutationCode[64] = { 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 3, 2, 2, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 2, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3};

   static int PermutationStackX[4][6] = { {0}, {0, 1}, {0, 0}, {0, 1, 0, 1, 0, 1}};
   static int PermutationStackY[4][6] = { {0}, {0, 2}, {0, 1}, {0, 2, 1, 2, 1, 2}};

   static int ColumnPermutations[64] = { 0, 1, 1, 5, 1, 3, 3, 11, 1, 3, 3, 11, 5, 11, 11, 35, 1, 3, 3, 11, 3, 7, 7, 23, 3, 7, 7, 23, 11, 23, 23, 71, 1, 3, 3, 11, 3, 7, 7, 23, 3, 7, 7, 23, 11, 23, 23, 71, 5, 11, 11, 35, 11, 23, 23, 71, 11, 23, 23, 71, 35, 71, 71, 215};
   static int PermutationColumnX[64][216] = {
   {0},
   {0, 7},
   {0, 6},
   {0, 7, 6, 7, 6, 7},
   {0, 4},
   {0, 7, 4, 7},
   {0, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7},
   {0, 3},
   {0, 7, 3, 7},
   {0, 6, 3, 6},
   {0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7},
   {0, 4, 3, 4, 3, 4},
   {0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7},
   {0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7},
   {0, 1},
   {0, 7, 1, 7},
   {0, 6, 1, 6},
   {0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7},
   {0, 4, 1, 4},
   {0, 7, 4, 7, 1, 7, 4, 7},
   {0, 6, 4, 6, 1, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7},
   {0, 3, 1, 3},
   {0, 7, 3, 7, 1, 7, 3, 7},
   {0, 6, 3, 6, 1, 6, 3, 6},
   {0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7},
   {0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4},
   {0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7},
   {0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7},
   {0, 0},
   {0, 7, 0, 7},
   {0, 6, 0, 6},
   {0, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7},
   {0, 4, 0, 4},
   {0, 7, 4, 7, 0, 7, 4, 7},
   {0, 6, 4, 6, 0, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7},
   {0, 3, 0, 3},
   {0, 7, 3, 7, 0, 7, 3, 7},
   {0, 6, 3, 6, 0, 6, 3, 6},
   {0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7},
   {0, 4, 3, 4, 3, 4, 0, 4, 3, 4, 3, 4},
   {0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7},
   {0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7},
   {0, 1, 0, 1, 0, 1},
   {0, 7, 1, 7, 0, 7, 1, 7, 0, 7, 1, 7},
   {0, 6, 1, 6, 0, 6, 1, 6, 0, 6, 1, 6},
   {0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7},
   {0, 4, 1, 4, 0, 4, 1, 4, 0, 4, 1, 4},
   {0, 7, 4, 7, 1, 7, 4, 7, 0, 7, 4, 7, 1, 7, 4, 7, 0, 7, 4, 7, 1, 7, 4, 7},
   {0, 6, 4, 6, 1, 6, 4, 6, 0, 6, 4, 6, 1, 6, 4, 6, 0, 6, 4, 6, 1, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7},
   {0, 3, 1, 3, 0, 3, 1, 3, 0, 3, 1, 3},
   {0, 7, 3, 7, 1, 7, 3, 7, 0, 7, 3, 7, 1, 7, 3, 7, 0, 7, 3, 7, 1, 7, 3, 7},
   {0, 6, 3, 6, 1, 6, 3, 6, 0, 6, 3, 6, 1, 6, 3, 6, 0, 6, 3, 6, 1, 6, 3, 6},
   {0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7},
   {0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4, 0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4, 0, 4, 3, 4, 3, 4, 1, 4, 3, 4, 3, 4},
   {0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 0, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7, 1, 7, 4, 7, 3, 7, 4, 7, 3, 7, 4, 7},
   {0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 0, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6, 1, 6, 4, 6, 3, 6, 4, 6, 3, 6, 4, 6},
   {0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 0, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 1, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7, 3, 7, 6, 7, 6, 7, 4, 7, 6, 7, 6, 7}
   };
   static int PermutationColumnY[64][216] = {
   {0},
   {0, 8},
   {0, 7},
   {0, 8, 7, 8, 7, 8},
   {0, 5},
   {0, 8, 5, 8},
   {0, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8},
   {0, 4},
   {0, 8, 4, 8},
   {0, 7, 4, 7},
   {0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8},
   {0, 5, 4, 5, 4, 5},
   {0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8},
   {0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8},
   {0, 2},
   {0, 8, 2, 8},
   {0, 7, 2, 7},
   {0, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8},
   {0, 5, 2, 5},
   {0, 8, 5, 8, 2, 8, 5, 8},
   {0, 7, 5, 7, 2, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8},
   {0, 4, 2, 4},
   {0, 8, 4, 8, 2, 8, 4, 8},
   {0, 7, 4, 7, 2, 7, 4, 7},
   {0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8},
   {0, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5},
   {0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8},
   {0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8},
   {0, 1},
   {0, 8, 1, 8},
   {0, 7, 1, 7},
   {0, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8},
   {0, 5, 1, 5},
   {0, 8, 5, 8, 1, 8, 5, 8},
   {0, 7, 5, 7, 1, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8},
   {0, 4, 1, 4},
   {0, 8, 4, 8, 1, 8, 4, 8},
   {0, 7, 4, 7, 1, 7, 4, 7},
   {0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8},
   {0, 5, 4, 5, 4, 5, 1, 5, 4, 5, 4, 5},
   {0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 1, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8},
   {0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 1, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8},
   {0, 2, 1, 2, 1, 2},
   {0, 8, 2, 8, 1, 8, 2, 8, 1, 8, 2, 8},
   {0, 7, 2, 7, 1, 7, 2, 7, 1, 7, 2, 7},
   {0, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8},
   {0, 5, 2, 5, 1, 5, 2, 5, 1, 5, 2, 5},
   {0, 8, 5, 8, 2, 8, 5, 8, 1, 8, 5, 8, 2, 8, 5, 8, 1, 8, 5, 8, 2, 8, 5, 8},
   {0, 7, 5, 7, 2, 7, 5, 7, 1, 7, 5, 7, 2, 7, 5, 7, 1, 7, 5, 7, 2, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8},
   {0, 4, 2, 4, 1, 4, 2, 4, 1, 4, 2, 4},
   {0, 8, 4, 8, 2, 8, 4, 8, 1, 8, 4, 8, 2, 8, 4, 8, 1, 8, 4, 8, 2, 8, 4, 8},
   {0, 7, 4, 7, 2, 7, 4, 7, 1, 7, 4, 7, 2, 7, 4, 7, 1, 7, 4, 7, 2, 7, 4, 7},
   {0, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8},
   {0, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5, 1, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5, 1, 5, 4, 5, 4, 5, 2, 5, 4, 5, 4, 5},
   {0, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 1, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 1, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8, 2, 8, 5, 8, 4, 8, 5, 8, 4, 8, 5, 8},
   {0, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 1, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 1, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7, 2, 7, 5, 7, 4, 7, 5, 7, 4, 7, 5, 7},
   {0, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 1, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 2, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8, 4, 8, 7, 8, 7, 8, 5, 8, 7, 8, 7, 8}
   };

   static int MiniRowOrderTrackerInit[18][3]
   {
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2},
      {0, 1, 2}
   };
   int MiniRowOrderTracker[18][3];

   int InputLineCount = InputBufferSize / 83;
   for (int inputgridix = 0; inputgridix < InputLineCount; inputgridix++)
   {
      for (int i=0; i<18; i++ ) { for (int j=0; j<3; j++ ) { OriginalMiniRowCount[ i][ j] = 0; } }
      for (int i=0; i<18; i++ ) { for (int j=0; j<3; j++ ) { DigitsInMiniRowBit[ i][ j] = 0; } }
      // Preliminary Step#1: Identify all row1 candidates based on minirow count pattern.
      // Count non-zero digits in minirows for direct and transposed puzzle.
      int CluesCount = 0;
      ArrayClear(InputGrid, 0, 81);
      int inputgridstart = inputgridix * 83;
         for (k = 0; k <= 80; k++) // Convert input char puzzle to integer array. Each non-zero (or non-period) character must be in range of 1 to 9.
         {
            char NextChr = InputGridBufferChr[inputgridstart + k];
            if (NextChr >= '1' && NextChr <= '9')
            {
                  InputGrid[k] = static_cast<int>(NextChr) - 48; // Convert Char to Integer.
                  CluesCount += 1;
                  i = k / 9; // i = Row - 0 to 8 notation
                  j = k % 9; // j = Column - 0 to 8 notation
                  m = j / 3; // m = MiniRow or Stack - 0 to 2 notation
                  n = i / 3; // n = MiniColumn or Band - 0 to 2 notation
                  OriginalMiniRowCount[i][m] += 1;
                  OriginalMiniRowCount[j + 9][n] += 1;
                  DigitsInMiniRowBit[i][m] |= IntToBit1To9[InputGrid[k]]; // Turn on digit's bit in direct row's minirow.
                  DigitsInMiniRowBit[j + 9][n] |= IntToBit1To9[InputGrid[k]]; // Turn on digit's bit in transposed row's minirow.
            }
         }
      if (CluesCount < 81) // Process sub grid (puzzles)
      {
         bool TranspositionNeededSw = false;
         ArrayConstrainedCopy(MinLexGridLocalReset, 0, MinLexGridLocal, 0, 81);
         // "Rightjustify" MiniRowCounts for rows of direct and transposed grids.
         //  NOTE: Isolated performance testing indicates that, for array size less than 20, a for loop is faster than Array.Clear() and Array.ConstrainedCopy().
         //        So they have been replaces by equivalent for loops for low array sizes. The origianl subroutine calls remain as comments.
         //Array.Clear(ZeroRowsInBandsCount, 0, 6)
         for (z = 0; z <= 5; z++)
         {
            ZeroRowsInBandsCount[z] = 0;
         }
         //Array.ConstrainedCopy(MinLexFirstNonZeroDigitPositionInRowReset, 4, MinLexFirstNonZeroDigitPositionInRow, 4, 6)
         for (z = 4; z <= 9; z++)
         {
            MinLexFirstNonZeroDigitPositionInRow[z] = MinLexFirstNonZeroDigitPositionInRowReset[z];
         }
         for (int i=0; i<18; i++ ) { for (int j=0; j<3; j++ ) { JustifiedMiniRowCount[ i][ j] = OriginalMiniRowCount[ i][ j]; } }
         for (int i=0; i<18; i++ ) { for (int j=0; j<3; j++ ) { MiniRowOrderTracker[ i][ j] = MiniRowOrderTrackerInit[ i][ j]; } }
         CalcMiniRowCountCodeMinimum = 99;
         for (i = 0; i <= 17; i++) // Right juastify direct and transposed band counts for eack row - high count to the right, e.g. 1,3,0 becomes 0,1,3.
         {
            DigitsInRowBit[i] = DigitsInMiniRowBit[i][0] | DigitsInMiniRowBit[i][1] | DigitsInMiniRowBit[i][2];
            if (DigitsInRowBit[i] == 0)
            {
               CalcMiniRowCountCode[i] = 0;
               ZeroRowsInBandsCount[i / 3] += 1;
            }
            else
            {
               if (JustifiedMiniRowCount[i][0] > JustifiedMiniRowCount[i][1]) // If minirow1 count is greater than minirow2 count, switch minirow1 and minirow2 counts.
               {
                  hold = JustifiedMiniRowCount[i][0];
                  JustifiedMiniRowCount[i][0] = JustifiedMiniRowCount[i][1];
                  JustifiedMiniRowCount[i][1] = hold;
                  MiniRowOrderTracker[i][0] = 1;
                  MiniRowOrderTracker[i][1] = 0;
               }
               if (JustifiedMiniRowCount[i][1] > JustifiedMiniRowCount[i][2]) // If minirow2 count is greater than MiniRow3 count, switch minirow2 and MiniRow3 counts.
               {
                  hold = JustifiedMiniRowCount[i][1];
                  JustifiedMiniRowCount[i][1] = JustifiedMiniRowCount[i][2];
                  JustifiedMiniRowCount[i][2] = hold;
                  hold = MiniRowOrderTracker[i][1];
                  MiniRowOrderTracker[i][1] = MiniRowOrderTracker[i][2];
                  MiniRowOrderTracker[i][2] = hold;
               }
               if (JustifiedMiniRowCount[i][0] > JustifiedMiniRowCount[i][1]) // If minirow1 count is greater than minirow2 count, switch minirow1 and minirow2 counts.
               {
                  hold = JustifiedMiniRowCount[i][0];
                  JustifiedMiniRowCount[i][0] = JustifiedMiniRowCount[i][1];
                  JustifiedMiniRowCount[i][1] = hold;
                  hold = MiniRowOrderTracker[i][0];
                  MiniRowOrderTracker[i][0] = MiniRowOrderTracker[i][1];
                  MiniRowOrderTracker[i][1] = hold;
               }
               CalcMiniRowCountCode[i] = 16 * JustifiedMiniRowCount[i][0] + 4 * JustifiedMiniRowCount[i][1] + JustifiedMiniRowCount[i][2];
            }
            if (CalcMiniRowCountCodeMinimum > CalcMiniRowCountCode[i])
            {
               CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode[i];
               zstart = i;
            }
         }

         CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = false;
         RowWithCalcMiniRowCountCodeMinimumIx = -1;
         for (i = zstart; i <= 17; i++) // Identify and save the rows that match the minimum count, these are candidates for row1.
         {
            //                                                                 (Note: All candidates will yield the same row1 after relabeling.)
            if (CalcMiniRowCountCodeMinimum == CalcMiniRowCountCode[i])
            {
               RowWithCalcMiniRowCountCodeMinimumIx += 1;
               RowWithCalcMiniRowCountCodeMinimum[RowWithCalcMiniRowCountCodeMinimumIx] = i; // NOTE: rows in 0 to 17 notation.
               if (!CandidatesRepeatDigitsBetweenFirstAndSecondRowSw && CalcMiniRowCountCodeMinimum > 0)
               {
                  RowRepeatDigits = DigitsInRowBit[i] & DigitsInRowBit[Row2CandidateA[i]];
                  if (RowRepeatDigits > 0)
                  {
                     CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = true;
                  }
                  else
                  {
                     RowRepeatDigits = DigitsInRowBit[i] & DigitsInRowBit[Row2CandidateB[i]];
                     if (RowRepeatDigits > 0)
                     {
                        CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = true;
                     }
                  }
               }
            }
         }

         FirstNonZeroPositionInRow1 = FirstNonZeroPositionInFirstNonZeroRow[CalcMiniRowCountCodeMinimum];
         //'Array.ConstrainedCopy(FirstNonZeroRowInit, CalcMiniRowCountCodeMinimum * 9, MinLexGridLocal, 0, 9) ' Set first row of MinLexGridLocal based on CalcMiniRowCountCodeMinimum.
         //zstart = CalcMiniRowCountCodeMinimum * 9
         //For z = 0 To 8 : MinLexGridLocal(z) = FirstNonZeroRowInit(z + zstart) : Next z
         TwoRowCandidateMiniRowCodeMinimum = 999;
         TwoRowCandidateMiniRow1CodeMinimum = 999;
         TwoRowCandidateIx = -1;
         Band1CandidateIx = -1;
         if (CalcMiniRowCountCodeMinimum > 0)
         {
            if (CalcMiniRowCountCodeMinimum < 4)
            {
               for (i = 0; i <= RowWithCalcMiniRowCountCodeMinimumIx; i++)
               {
                  TwoRowCandidateIx += 1;
                  j = RowWithCalcMiniRowCountCodeMinimum[i];
                  k = Row2CandidateA[j];
                  m = MiniRowOrderTracker[j][0];
                  n = MiniRowOrderTracker[j][1];
                  if (CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][m];
                     if (MiniRowRepeatDigits > 0)
                     {
                        AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m] * 2 - 1;
                     }
                     else
                     {
                        AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m] * 2;
                     }
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][n];
                     if (MiniRowRepeatDigits > 0)
                     {
                        AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n] * 2 - 1;
                     }
                     else
                     {
                        AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n] * 2;
                     }
                  }
                  else
                  {
                     AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m];
                     AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n];
                  }
                  if (AugmentedMiniRowCount1 > AugmentedMiniRowCount2)
                  {
                     TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount2;
                     TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount1;
                  }
                  else
                  {
                     TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount1;
                     TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount2;
                  }
                  TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = TwoRowCandidateMiniRow1Count * 10 + TwoRowCandidateMiniRow2Count;
                  if (TwoRowCandidateMiniRowCodeMinimum > TwoRowCandidateMiniRowCode[TwoRowCandidateIx])
                  {
                     TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode[TwoRowCandidateIx];
                     zstart = i;
                  }
                  TwoRowCandidateRow1[TwoRowCandidateIx] = j;
                  TwoRowCandidateRow2[TwoRowCandidateIx] = k;
                  TwoRowCandidateRow3[TwoRowCandidateIx] = Row2CandidateB[j];
                  TwoRowCandidateIx += 1;
                  k = Row2CandidateB[j];
                  if (CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][m];
                     if (MiniRowRepeatDigits > 0)
                     {
                        AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m] * 2 - 1;
                     }
                     else
                     {
                        AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m] * 2;
                     }
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][n];
                     if (MiniRowRepeatDigits > 0)
                     {
                        AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n] * 2 - 1;
                     }
                     else
                     {
                        AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n] * 2;
                     }
                  }
                  else
                  {
                     AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m];
                     AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n];
                  }
                  if (AugmentedMiniRowCount1 > AugmentedMiniRowCount2)
                  {
                     TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount2;
                     TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount1;
                  }
                  else
                  {
                     TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount1;
                     TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount2;
                  }
                  TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = TwoRowCandidateMiniRow1Count * 10 + TwoRowCandidateMiniRow2Count;
                  if (TwoRowCandidateMiniRowCodeMinimum > TwoRowCandidateMiniRowCode[TwoRowCandidateIx])
                  {
                     TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode[TwoRowCandidateIx];
                     zstart = i;
                  }
                  TwoRowCandidateRow1[TwoRowCandidateIx] = j;
                  TwoRowCandidateRow2[TwoRowCandidateIx] = k;
                  TwoRowCandidateRow3[TwoRowCandidateIx] = Row2CandidateA[j];
               }
               for (i = zstart; i <= TwoRowCandidateIx; i++)
               {
                  if (TwoRowCandidateMiniRowCodeMinimum == TwoRowCandidateMiniRowCode[i])
                  {
                     Band1CandidateIx += 1;
                     Band1CandidateRow1[Band1CandidateIx] = TwoRowCandidateRow1[i];
                     Band1CandidateRow2[Band1CandidateIx] = TwoRowCandidateRow2[i];
                     Band1CandidateRow3[Band1CandidateIx] = TwoRowCandidateRow3[i];
                     if (Band1CandidateRow1[Band1CandidateIx] > 8)
                     {
                        TranspositionNeededSw = true;
                     }
                  }
               }
            }
            else if (CalcMiniRowCountCodeMinimum < 16)
            {
               for (i = 0; i <= RowWithCalcMiniRowCountCodeMinimumIx; i++)
               {
                  TwoRowCandidateIx += 1;
                  j = RowWithCalcMiniRowCountCodeMinimum[i];
                  k = Row2CandidateA[j];
                  m = MiniRowOrderTracker[j][0];
                  if (CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][m];
                     if (MiniRowRepeatDigits > 0)
                     {
                        TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = OriginalMiniRowCount[k][m] * 2 - 1;
                     }
                     else
                     {
                        TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = OriginalMiniRowCount[k][m] * 2;
                     }
                  }
                  else
                  {
                     TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = OriginalMiniRowCount[k][m];
                  }
                  if (TwoRowCandidateMiniRow1CodeMinimum > TwoRowCandidateMiniRowCode[TwoRowCandidateIx])
                  {
                     TwoRowCandidateMiniRow1CodeMinimum = TwoRowCandidateMiniRowCode[TwoRowCandidateIx];
                     zstart = TwoRowCandidateIx;
                  }
                  TwoRowCandidateRow1[TwoRowCandidateIx] = j;
                  TwoRowCandidateRow2[TwoRowCandidateIx] = k;
                  TwoRowCandidateRow3[TwoRowCandidateIx] = Row2CandidateB[j];
                  TwoRowCandidateIx += 1;
                  k = Row2CandidateB[j];
                  if (CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][m];
                     if (MiniRowRepeatDigits > 0)
                     {
                        TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = OriginalMiniRowCount[k][m] * 2 - 1;
                     }
                     else
                     {
                        TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = OriginalMiniRowCount[k][m] * 2;
                     }
                  }
                  else
                  {
                     TwoRowCandidateMiniRowCode[TwoRowCandidateIx] = OriginalMiniRowCount[k][m];
                  }
                  if (TwoRowCandidateMiniRow1CodeMinimum > TwoRowCandidateMiniRowCode[TwoRowCandidateIx])
                  {
                     TwoRowCandidateMiniRow1CodeMinimum = TwoRowCandidateMiniRowCode[TwoRowCandidateIx];
                     zstart = TwoRowCandidateIx;
                  }
                  TwoRowCandidateRow1[TwoRowCandidateIx] = j;
                  TwoRowCandidateRow2[TwoRowCandidateIx] = k;
                  TwoRowCandidateRow3[TwoRowCandidateIx] = Row2CandidateA[j];
               }
               for (i = zstart; i <= TwoRowCandidateIx; i++)
               {
                  if (TwoRowCandidateMiniRow1CodeMinimum == TwoRowCandidateMiniRowCode[i])
                  {
                     Band1CandidateIx += 1;
                     Band1CandidateRow1[Band1CandidateIx] = TwoRowCandidateRow1[i];
                     Band1CandidateRow2[Band1CandidateIx] = TwoRowCandidateRow2[i];
                     Band1CandidateRow3[Band1CandidateIx] = TwoRowCandidateRow3[i];
                     if (Band1CandidateRow1[Band1CandidateIx] > 8)
                     {
                        TranspositionNeededSw = true;
                     }
                  }
               }
            }
            else // CalcMiniRowCountCodeMinimum > 15
            {
               for (i = 0; i <= RowWithCalcMiniRowCountCodeMinimumIx; i++)
               {
                  z = RowWithCalcMiniRowCountCodeMinimum[i];
                  a = Row2CandidateA[z];
                  b = Row2CandidateB[z];
                  Band1CandidateIx += 1;
                  Band1CandidateRow1[Band1CandidateIx] = z;
                  Band1CandidateRow2[Band1CandidateIx] = a;
                  Band1CandidateRow3[Band1CandidateIx] = b;
                  if (z > 8)
                  {
                     TranspositionNeededSw = true;
                  }
                  Band1CandidateIx += 1;
                  Band1CandidateRow1[Band1CandidateIx] = z;
                  Band1CandidateRow2[Band1CandidateIx] = b;
                  Band1CandidateRow3[Band1CandidateIx] = a;
               }
            }
         }
         else // If rows with all zeros exist, possibly reduce band 1 candidates by evaluating their row2 candidates.
         {
            // Identify and save the rows that match the row2 minimum count for all-zeros row1 candidates, these are candidates for row2.
            // (Note: All row2 candidates will yield the same row2 after relabeling.)
            CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = false;
            Row2CalcMiniRowCountCodeMinimum = 999;
            TwoRowCandidateIx = -1;
            for (i = 0; i <= RowWithCalcMiniRowCountCodeMinimumIx; i++)
            {
               z = RowWithCalcMiniRowCountCodeMinimum[i];
               a = Row2CandidateA[z];
               b = Row2CandidateB[z];
               if (Row2CalcMiniRowCountCodeMinimum > CalcMiniRowCountCode[a])
               {
                  Row2CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode[a];
                  zstart = i;
               }
               if (Row2CalcMiniRowCountCodeMinimum > CalcMiniRowCountCode[b])
               {
                  Row2CalcMiniRowCountCodeMinimum = CalcMiniRowCountCode[b];
                  zstart = i;
               }
            }
            for (i = zstart; i <= RowWithCalcMiniRowCountCodeMinimumIx; i++)
            {
               z = RowWithCalcMiniRowCountCodeMinimum[i];
               a = Row2CandidateA[z];
               b = Row2CandidateB[z];
               if (Row2CalcMiniRowCountCodeMinimum == CalcMiniRowCountCode[a])
               {
                  TwoRowCandidateIx += 1;
                  TwoRowCandidateRow1[TwoRowCandidateIx] = z;
                  TwoRowCandidateRow2[TwoRowCandidateIx] = a;
                  TwoRowCandidateRow3[TwoRowCandidateIx] = b;
                  if (!CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     RowRepeatDigits = DigitsInRowBit[a] & DigitsInRowBit[b];
                     if (RowRepeatDigits > 0)
                     {
                        CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = true;
                     }
                  }
               }
               if (Row2CalcMiniRowCountCodeMinimum == CalcMiniRowCountCode[b])
               {
                  TwoRowCandidateIx += 1;
                  TwoRowCandidateRow1[TwoRowCandidateIx] = z;
                  TwoRowCandidateRow2[TwoRowCandidateIx] = b;
                  TwoRowCandidateRow3[TwoRowCandidateIx] = a;
                  if (!CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     RowRepeatDigits = DigitsInRowBit[b] & DigitsInRowBit[a];
                     if (RowRepeatDigits > 0)
                     {
                        CandidatesRepeatDigitsBetweenFirstAndSecondRowSw = true;
                     }
                  }
               }
            }
            if (TwoRowCandidateIx == 0)
            {
               Band1CandidateIx = 0;
               Band1CandidateRow1[0] = TwoRowCandidateRow1[0];
               Band1CandidateRow2[0] = TwoRowCandidateRow2[0];
               Band1CandidateRow3[0] = TwoRowCandidateRow3[0];
               if (Band1CandidateRow1[Band1CandidateIx] > 8)
               {
                  TranspositionNeededSw = true;
               }
            }
            else if (Row2CalcMiniRowCountCodeMinimum < 4)
            {
               for (i = 0; i <= TwoRowCandidateIx; i++)
               {
                  j = TwoRowCandidateRow2[i];
                  k = TwoRowCandidateRow3[i];
                  m = MiniRowOrderTracker[j][0];
                  n = MiniRowOrderTracker[j][1];
                  if (CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][m];
                     if (MiniRowRepeatDigits > 0)
                     {
                        AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m] * 2 - 1;
                     }
                     else
                     {
                        AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m] * 2;
                     }
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][n];
                     if (MiniRowRepeatDigits > 0)
                     {
                        AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n] * 2 - 1;
                     }
                     else
                     {
                        AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n] * 2;
                     }
                  }
                  else
                  {
                     AugmentedMiniRowCount1 = OriginalMiniRowCount[k][m];
                     AugmentedMiniRowCount2 = OriginalMiniRowCount[k][n];
                  }
                  if (AugmentedMiniRowCount1 > AugmentedMiniRowCount2)
                  {
                     TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount2;
                     TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount1;
                  }
                  else
                  {
                     TwoRowCandidateMiniRow1Count = AugmentedMiniRowCount1;
                     TwoRowCandidateMiniRow2Count = AugmentedMiniRowCount2;
                  }
                  TwoRowCandidateMiniRowCode[i] = TwoRowCandidateMiniRow1Count * 10 + TwoRowCandidateMiniRow2Count;
                  if (TwoRowCandidateMiniRowCodeMinimum > TwoRowCandidateMiniRowCode[i])
                  {
                     TwoRowCandidateMiniRowCodeMinimum = TwoRowCandidateMiniRowCode[i];
                     zstart = i;
                  }
               }
               for (i = zstart; i <= TwoRowCandidateIx; i++)
               {
                  if (TwoRowCandidateMiniRowCodeMinimum == TwoRowCandidateMiniRowCode[i])
                  {
                     Band1CandidateIx += 1;
                     Band1CandidateRow1[Band1CandidateIx] = TwoRowCandidateRow1[i];
                     Band1CandidateRow2[Band1CandidateIx] = TwoRowCandidateRow2[i];
                     Band1CandidateRow3[Band1CandidateIx] = TwoRowCandidateRow3[i];
                     if (Band1CandidateRow1[Band1CandidateIx] > 8)
                     {
                        TranspositionNeededSw = true;
                     }
                  }
               }
            }
            else if (Row2CalcMiniRowCountCodeMinimum < 16)
            {
               for (i = 0; i <= TwoRowCandidateIx; i++)
               {
                  j = TwoRowCandidateRow2[i];
                  k = TwoRowCandidateRow3[i];
                  m = MiniRowOrderTracker[j][0];
                  if (CandidatesRepeatDigitsBetweenFirstAndSecondRowSw)
                  {
                     MiniRowRepeatDigits = DigitsInRowBit[j] & DigitsInMiniRowBit[k][m];
                     if (MiniRowRepeatDigits > 0)
                     {
                        TwoRowCandidateMiniRowCode[i] = OriginalMiniRowCount[k][m] * 2 - 1;
                     }
                     else
                     {
                        TwoRowCandidateMiniRowCode[i] = OriginalMiniRowCount[k][m] * 2;
                     }
                  }
                  else
                  {
                     TwoRowCandidateMiniRowCode[i] = OriginalMiniRowCount[k][m];
                  }
                  if (TwoRowCandidateMiniRow1CodeMinimum > TwoRowCandidateMiniRowCode[i])
                  {
                     TwoRowCandidateMiniRow1CodeMinimum = TwoRowCandidateMiniRowCode[i];
                     zstart = i;
                  }
               }
               for (i = zstart; i <= TwoRowCandidateIx; i++)
               {
                  if (TwoRowCandidateMiniRow1CodeMinimum == TwoRowCandidateMiniRowCode[i])
                  {
                     Band1CandidateIx += 1;
                     Band1CandidateRow1[Band1CandidateIx] = TwoRowCandidateRow1[i];
                     Band1CandidateRow2[Band1CandidateIx] = TwoRowCandidateRow2[i];
                     Band1CandidateRow3[Band1CandidateIx] = TwoRowCandidateRow3[i];
                     if (Band1CandidateRow1[Band1CandidateIx] > 8)
                     {
                        TranspositionNeededSw = true;
                     }
                  }
               }
            }
            else // Row2CalcMiniRowCountCodeMinimum > 15
            {
               for (i = 0; i <= TwoRowCandidateIx; i++)
               {
                  Band1CandidateIx += 1;
                  Band1CandidateRow1[Band1CandidateIx] = TwoRowCandidateRow1[i];
                  Band1CandidateRow2[Band1CandidateIx] = TwoRowCandidateRow2[i];
                  Band1CandidateRow3[Band1CandidateIx] = TwoRowCandidateRow3[i];
                  if (TwoRowCandidateRow1[i] > 8)
                  {
                     TranspositionNeededSw = true;
                  }
               }
            }
            FirstNonZeroPositionInRow2 = FirstNonZeroPositionInFirstNonZeroRow[Row2CalcMiniRowCountCodeMinimum];
            //'Array.ConstrainedCopy(FirstNonZeroRowInit, Row2CalcMiniRowCountCodeMinimum * 9, MinLexGridLocal, 9, 9) ' For cases of all zeros row1, Set the second row of MinLexGridLocal based on Row2CalcMiniRowCountCodeMinimum.
            //zstart = Row2CalcMiniRowCountCodeMinimum * 9
            //For z = 0 To 8 : MinLexGridLocal(z + 9) = FirstNonZeroRowInit(z + zstart) : Next z
         }

         if (TranspositionNeededSw)
         {
            for (i = 0; i <= 80; i++) // transpose rows to columns.
            {
               InputGrid[81 + i / 9 + (i % 9) * 9] = InputGrid[i];
            }
         }

         // Band1 Processing - Two cases:
         //    1) MinLexBand1:     If all 18 rows and columns have at least one non-zero digit, then MinLex Band1 for all row1 candidates and eliminate Band1 row/column, permutation & relabel possibilities that do not yield the Band1 MinLex.
         //    2) MinLexRows2and3: If all-zeros rows or columns exist, then MinLex Band1 for all row2 candidates for the all-zeros row and columns and eliminate Band1 row/column, permutation & relabel possibilities that do not yield the Band1 MinLex.
         //    Results:            1) MinLexed Band1 solution candidates ready for rows 4 - 9 processing.
         //                        2) Column (and stack) permutation Trackers used to reproduce each puzzle configuration that yields its Band1 MinLex.
         //                        3) Relabel Trackers for each Band1 MinLex configuration.

         StoppedJustifyingRow = 0;
         if (CalcMiniRowCountCodeMinimum > 0) // Start MinLexBand1 (Row1 not empty case)
         {
            FirstNonZeroDigitPositionInRow[2] = -1;
            Row1StackPermutationCode = FirstNonZeroRowStackPermutationCode[CalcMiniRowCountCodeMinimum];
            ArrayConstrainedCopy(MinLexGridLocal, 0, MinLexRows, 0, 27); // Initialize MinLexRows (Band1 MinLex work array) to all 10s.
            ColumnPermutationTrackerIx = -1;
            for (int candidate = 0; candidate <= Band1CandidateIx; candidate++)
            {
               Row1Candidate17 = Band1CandidateRow1[candidate];
               if (Row1Candidate17 < 9)
               {
                  Row2Candidate17 = Band1CandidateRow2[candidate];
                  Row1Candidate = Row1Candidate17 + 1;
                  Row2Candidate = Band1CandidateRow2[candidate] + 1;
                  Row3Candidate = Band1CandidateRow3[candidate] + 1;
                  ArrayConstrainedCopy(InputGrid, 0, HoldPuzzle, 0, 81); // Copy the direct input puzzle to HoldPuzzle.
               }
               else
               {
                  Row2Candidate17 = Band1CandidateRow2[candidate];
                  Row1Candidate = Row1Candidate17 - 8;
                  Row2Candidate = Band1CandidateRow2[candidate] - 8;
                  Row3Candidate = Band1CandidateRow3[candidate] - 8;
                  ArrayConstrainedCopy(InputGrid, 81, HoldPuzzle, 0, 81); // Copy the transposed input puzzle to HoldPuzzle.
               }
               // MinLex Band1 for Row1Candidate.
               MinLexRowHitSw = false;
               ResetMinLexRowSw = false;
               CandidateColumnPermutationTrackerStartIx = ColumnPermutationTrackerIx;
               row1start = (Row1Candidate - 1) * 9;
               row2start = (Row2Candidate - 1) * 9;
               row3start = (Row3Candidate - 1) * 9;

               StackPermutationTrackerLocal[0] = MiniRowOrderTracker[Row1Candidate17][0];
               StackPermutationTrackerLocal[1] = MiniRowOrderTracker[Row1Candidate17][1];
               StackPermutationTrackerLocal[2] = MiniRowOrderTracker[Row1Candidate17][2];
               Row1MiniRowCount[0] = JustifiedMiniRowCount[Row1Candidate17][0];
               Row1MiniRowCount[1] = JustifiedMiniRowCount[Row1Candidate17][1];
               Row1MiniRowCount[2] = JustifiedMiniRowCount[Row1Candidate17][2];
               Row2MiniRowCount[0] = OriginalMiniRowCount[Row2Candidate17][StackPermutationTrackerLocal[0]];
               Row2MiniRowCount[1] = OriginalMiniRowCount[Row2Candidate17][StackPermutationTrackerLocal[1]];
               Row2MiniRowCount[2] = OriginalMiniRowCount[Row2Candidate17][StackPermutationTrackerLocal[2]];

               Row2Or3StackPermutationCode = Row1StackPermutationCode;
               if (CalcMiniRowCountCodeMinimum < 4 || Row1StackPermutationCode > 1)
               {
                  if (Row2MiniRowCount[0] > Row2MiniRowCount[1]) // If Row2 minirow1 count is greater than minirow2, switch minirow1 and minirow2.
                  {
                     hold = Row2MiniRowCount[0];
                     Row2MiniRowCount[0] = Row2MiniRowCount[1];
                     Row2MiniRowCount[1] = hold;
                     hold = StackPermutationTrackerLocal[0];
                     StackPermutationTrackerLocal[0] = StackPermutationTrackerLocal[1];
                     StackPermutationTrackerLocal[1] = hold;
                  }
                  else if (CalcMiniRowCountCodeMinimum < 4 && Row2MiniRowCount[1] > 0 && Row2MiniRowCount[0] == Row2MiniRowCount[1])
                  {
                     Row2Or3StackPermutationCode = 2;
                  }
               }

               MinLexCandidateSw = true;
               if (Row1StackPermutationCode == 0)
               {
                  //Select Case FirstNonZeroDigitPositionInRow(2)
                  //Case < 1 // If -1 or 0, do nothing.
                  if (FirstNonZeroDigitPositionInRow[2] < 1) // If -1 or 0, do nothing.
                  {
                  }
                  //Case 5
                  else if (FirstNonZeroDigitPositionInRow[2] == 5)
                  {
                     if (Row2MiniRowCount[0] > 0 || Row2MiniRowCount[1] > 1)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 4
                  else if (FirstNonZeroDigitPositionInRow[2] == 4)
                  {
                     if (Row2MiniRowCount[0] > 0 || Row2MiniRowCount[1] == 3)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case > 5
                  else if (FirstNonZeroDigitPositionInRow[2] > 5)
                  {
                     if (Row2MiniRowCount[0] > 0 || Row2MiniRowCount[1] > 0)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 3
                  else if (FirstNonZeroDigitPositionInRow[2] == 3)
                  {
                     if (Row2MiniRowCount[0] > 0)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 2
                  else if (FirstNonZeroDigitPositionInRow[2] == 2)
                  {
                     if (Row2MiniRowCount[0] > 1)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 1
                  else if (FirstNonZeroDigitPositionInRow[2] == 1)
                  {
                     if (Row2MiniRowCount[0] == 3)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
               }

               if (MinLexCandidateSw)
               {
                  StackPermutationTrackerCode = StackPermutationTrackerLocal[0] * 100 + StackPermutationTrackerLocal[1] * 10 + StackPermutationTrackerLocal[2];
                  switch (StackPermutationTrackerCode) // Apply permutations required to right justify row 1 and row 2.
                  {
                     case 12: // 012 - no permutation
                     break;
                     case 21: // 021
                        SwitchStacks12(HoldPuzzle);
                        break;
                     case 102:
                        SwitchStacks01(HoldPuzzle);
                        break;
                     case 120:
                        Switch3Stacks120(HoldPuzzle);
                        break;
                     case 201:
                        Switch3Stacks201(HoldPuzzle);
                        break;
                     case 210:
                        SwitchStacks02(HoldPuzzle);
                        break;
                  }

                  Row3MiniRowCount[0] = 0;
                  Row3MiniRowCount[1] = 0;
                  Row3MiniRowCount[2] = 0;
                  for (i = 0; i <= 8; i++) // Count row 3 non-zero digits in MiniRows 1, 2 and 3.
                  {
                     if (HoldPuzzle[row3start + i] > 0)
                     {
                        Row3MiniRowCount[i / 3] += 1;
                     }
                  }
                  if (Row2MiniRowCount[0] == 0 && Row2MiniRowCount[1] == 0)
                  {
                     if (Row3MiniRowCount[0] > Row3MiniRowCount[1]) // If row 3 minirow1 count is greater than minirow2, switch minirow1 and minirow2.
                     {
                        hold = Row3MiniRowCount[0];
                        Row3MiniRowCount[0] = Row3MiniRowCount[1];
                        Row3MiniRowCount[1] = hold;
                        SwitchStacks01(HoldPuzzle);
                     }
                     else if (Row3MiniRowCount[1] > 0 && Row3MiniRowCount[0] == Row3MiniRowCount[1])
                     {
                        Row2Or3StackPermutationCode = 2;
                     }
                  }

                  // Positionally (not considering digit values) "right justify" first, second and third row non-zero digits within MiniRows.
                  // Also set the ColumnPermutationCode for Band1 using MiniRow 0 to 2 notation.
                  Band1MiniRowColumnPermutationCode[0] = 0;
                  Band1MiniRowColumnPermutationCode[1] = 0;
                  Band1MiniRowColumnPermutationCode[2] = 0;
                  switch (Row1MiniRowCount[0]) // MiniRow 1
                  {
                     case 0: // Row1MiniRowCount(0) = 0
                        switch (Row2MiniRowCount[0])
                        {
                           case 0: // Row 2
                              switch (Row3MiniRowCount[0])
                              {
                                 case 0: // Do nothing.
                                 break;
                                 case 1:
                                    if (HoldPuzzle[row3start] > 0)
                                    {
                                       SwitchColumns02(HoldPuzzle);
                                    }
                                    else if (HoldPuzzle[row3start + 1] > 0)
                                    {
                                       SwitchColumns12(HoldPuzzle);
                                    }
                                    break;
                                 case 2:
                                    if (HoldPuzzle[row3start + 2] == 0)
                                    {
                                       SwitchColumns02(HoldPuzzle);
                                    }
                                    else if (HoldPuzzle[row3start + 1] == 0)
                                    {
                                       SwitchColumns01(HoldPuzzle);
                                    }
                                    Band1MiniRowColumnPermutationCode[0] += 1;
                                    break;
                                 case 3:
                                    Band1MiniRowColumnPermutationCode[0] += 3;
                                    break;
                              }
                              break;
                           case 1: // Row2
                              if (HoldPuzzle[row2start] > 0)
                              {
                                 SwitchColumns02(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row2start + 1] > 0)
                              {
                                 SwitchColumns12(HoldPuzzle);
                              }
                              if (HoldPuzzle[row3start] > 0)
                              {
                                 if (HoldPuzzle[row3start + 1] == 0)
                                 {
                                    SwitchColumns01(HoldPuzzle);
                                 }
                                 else
                                 {
                                    Band1MiniRowColumnPermutationCode[0] += 2;
                                 }
                              }
                              break;
                           case 2: // Row2
                              if (HoldPuzzle[row2start + 2] == 0)
                              {
                                 SwitchColumns02(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row2start + 1] == 0)
                              {
                                 SwitchColumns01(HoldPuzzle);
                              }
                              Band1MiniRowColumnPermutationCode[0] += 1;
                              break;
                           case 3: // Row2
                              Band1MiniRowColumnPermutationCode[0] += 3;
                              break;
                        }
                        break;
                     case 1: // Row1MiniRowCount(0) = 1
                        if (HoldPuzzle[row1start] > 0)
                        {
                           SwitchColumns02(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row1start + 1] > 0)
                        {
                           SwitchColumns12(HoldPuzzle);
                        }
                        if (HoldPuzzle[row2start] > 0)
                        {
                           if (HoldPuzzle[row2start + 1] == 0)
                           {
                              SwitchColumns01(HoldPuzzle);
                           }
                           else
                           {
                              Band1MiniRowColumnPermutationCode[0] += 2;
                           }
                        }
                        else
                        {
                           if (HoldPuzzle[row2start + 1] == 0)
                           {
                              if (HoldPuzzle[row3start] > 0)
                              {
                                 if (HoldPuzzle[row3start + 1] == 0)
                                 {
                                    SwitchColumns01(HoldPuzzle);
                                 }
                                 else
                                 {
                                    Band1MiniRowColumnPermutationCode[0] += 2;
                                 }
                              }
                           }
                        }
                        break;
                     case 2: // Row1MiniRowCount(0) = 2
                        if (HoldPuzzle[row1start + 2] == 0)
                        {
                           SwitchColumns02(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row1start + 1] == 0)
                        {
                           SwitchColumns01(HoldPuzzle);
                        }
                        Band1MiniRowColumnPermutationCode[0] += 1;
                        break;
                     case 3: // Row1MiniRowCount(0) = 3
                        Band1MiniRowColumnPermutationCode[0] += 3;
                        break;
                  }

                  switch (Row1MiniRowCount[1]) // MiniRow 2
                  {
                     case 0: // Row1MiniRowCount(1) = 0
                        switch (Row2MiniRowCount[1])
                        {
                           case 0: // Row 2
                              switch (Row3MiniRowCount[1])
                              {
                                 case 0: // Do nothing.
                                 break;
                                 case 1:
                                    if (HoldPuzzle[row3start + 3] > 0)
                                    {
                                       SwitchColumns35(HoldPuzzle);
                                    }
                                    else if (HoldPuzzle[row3start + 4] > 0)
                                    {
                                       SwitchColumns45(HoldPuzzle);
                                    }
                                    break;
                                 case 2:
                                    if (HoldPuzzle[row3start + 5] == 0)
                                    {
                                       SwitchColumns35(HoldPuzzle);
                                    }
                                    else if (HoldPuzzle[row3start + 4] == 0)
                                    {
                                       SwitchColumns34(HoldPuzzle);
                                    }
                                    Band1MiniRowColumnPermutationCode[1] += 1;
                                    break;
                                 case 3:
                                    Band1MiniRowColumnPermutationCode[1] += 3;
                                    break;
                              }
                              break;
                           case 1: // Row2
                              if (HoldPuzzle[row2start + 3] > 0)
                              {
                                 SwitchColumns35(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row2start + 4] > 0)
                              {
                                 SwitchColumns45(HoldPuzzle);
                              }
                              if (HoldPuzzle[row3start + 3] > 0)
                              {
                                 if (HoldPuzzle[row3start + 4] == 0)
                                 {
                                    SwitchColumns34(HoldPuzzle);
                                 }
                                 else
                                 {
                                    Band1MiniRowColumnPermutationCode[1] += 2;
                                 }
                              }
                              break;
                           case 2: // Row2
                              if (HoldPuzzle[row2start + 5] == 0)
                              {
                                 SwitchColumns35(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row2start + 4] == 0)
                              {
                                 SwitchColumns34(HoldPuzzle);
                              }
                              Band1MiniRowColumnPermutationCode[1] += 1;
                              break;
                           case 3: // Row2
                              Band1MiniRowColumnPermutationCode[1] += 3;
                              break;
                        }
                        break;
                     case 1: // Row1MiniRowCount(1) = 1
                        if (HoldPuzzle[row1start + 3] > 0)
                        {
                           SwitchColumns35(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row1start + 4] > 0)
                        {
                           SwitchColumns45(HoldPuzzle);
                        }
                        if (HoldPuzzle[row2start + 3] > 0)
                        {
                           if (HoldPuzzle[row2start + 4] == 0)
                           {
                              SwitchColumns34(HoldPuzzle);
                           }
                           else
                           {
                              Band1MiniRowColumnPermutationCode[1] += 2;
                           }
                        }
                        else
                        {
                           if (HoldPuzzle[row2start + 4] == 0)
                           {
                              if (HoldPuzzle[row3start + 3] > 0)
                              {
                                 if (HoldPuzzle[row3start + 4] == 0)
                                 {
                                    SwitchColumns34(HoldPuzzle);
                                 }
                                 else
                                 {
                                    Band1MiniRowColumnPermutationCode[1] += 2;
                                 }
                              }
                           }
                        }
                        break;
                     case 2: // Row1MiniRowCount(1) = 2
                        if (HoldPuzzle[row1start + 5] == 0)
                        {
                           SwitchColumns35(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row1start + 4] == 0)
                        {
                           SwitchColumns34(HoldPuzzle);
                        }
                        Band1MiniRowColumnPermutationCode[1] += 1;
                        break;
                     case 3: // Row1MiniRowCount(1) = 3
                        Band1MiniRowColumnPermutationCode[1] += 3;
                        break;
                  }

                  switch (Row1MiniRowCount[2]) // MiniRow 3
                  {
                  // Case 0 not possible since this section handles non-zero row1 candidates.
                     case 1: // Row1MiniRowCount(2) = 1
                        if (HoldPuzzle[row1start + 6] > 0)
                        {
                           SwitchColumns68(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row1start + 7] > 0)
                        {
                           SwitchColumns78(HoldPuzzle);
                        }
                        if (HoldPuzzle[row2start + 6] > 0)
                        {
                           if (HoldPuzzle[row2start + 7] == 0)
                           {
                              SwitchColumns67(HoldPuzzle);
                           }
                           else
                           {
                              Band1MiniRowColumnPermutationCode[2] += 2;
                           }
                        }
                        else
                        {
                           if (HoldPuzzle[row2start + 7] == 0)
                           {
                              if (HoldPuzzle[row3start + 6] > 0)
                              {
                                 if (HoldPuzzle[row3start + 7] == 0)
                                 {
                                    SwitchColumns67(HoldPuzzle);
                                 }
                                 else
                                 {
                                    Band1MiniRowColumnPermutationCode[2] += 2;
                                 }
                              }
                           }
                        }
                        break;
                     case 2: // Row1MiniRowCount(2) = 2
                        if (HoldPuzzle[row1start + 8] == 0)
                        {
                           SwitchColumns68(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row1start + 7] == 0)
                        {
                           SwitchColumns67(HoldPuzzle);
                        }
                        Band1MiniRowColumnPermutationCode[2] += 1;
                        break;
                     case 3: // Row1MiniRowCount(2) = 3
                        Band1MiniRowColumnPermutationCode[2] += 3;
                        break;
                  }

                  //Array.ConstrainedCopy(HoldPuzzle, (Row1Candidate - 1) * 9, LocalBand1, 0, 9)
                  zstart = (Row1Candidate - 1) * 9;
                  for (z = 0; z <= 8; z++)
                  {
                     LocalBand1[z] = HoldPuzzle[z + zstart];
                  }
                  //Array.ConstrainedCopy(HoldPuzzle, (Row2Candidate - 1) * 9, LocalBand1, 9, 9)
                  zstart = (Row2Candidate - 1) * 9;
                  for (z = 0; z <= 8; z++)
                  {
                     LocalBand1[z + 9] = HoldPuzzle[z + zstart];
                  }
                  //Array.ConstrainedCopy(HoldPuzzle, (Row3Candidate - 1) * 9, LocalBand1, 18, 9)
                  zstart = (Row3Candidate - 1) * 9;
                  for (z = 0; z <= 8; z++)
                  {
                     LocalBand1[z + 18] = HoldPuzzle[z + zstart];
                  }
                  //Array.ConstrainedCopy(ColumnTrackerInit, 0, LocalColumnPermutationTracker, 0, 9)
                  for (z = 0; z <= 8; z++)
                  {
                     LocalColumnPermutationTracker[z] = ColumnTrackerInit[z];
                  }
                  FirstColumnPermutationTrackerIsIdentitySw = true;
                  int tempVar1 = StackPermutations[ Row2Or3StackPermutationCode];
                  for (int stackpermutationix = 0; stackpermutationix <= tempVar1; stackpermutationix++)
                  {
                     if (stackpermutationix > 0)
                     {
                        FirstColumnPermutationTrackerIsIdentitySw = false;
                        stackx = PermutationStackX[Row2Or3StackPermutationCode][stackpermutationix]; // Band1 (3 row) stack switch.
                        stacky = PermutationStackY[Row2Or3StackPermutationCode][stackpermutationix];
                        switchx = 3 * stackx;
                        switchy = 3 * stacky;
                        hold = Band1MiniRowColumnPermutationCode[stackx];
                        Band1MiniRowColumnPermutationCode[stackx] = Band1MiniRowColumnPermutationCode[stacky];
                        Band1MiniRowColumnPermutationCode[stacky] = hold;
                        hold = Row2MiniRowCount[stackx];
                        Row2MiniRowCount[stackx] = Row2MiniRowCount[stacky];
                        Row2MiniRowCount[stacky] = hold;
                        hold = LocalColumnPermutationTracker[switchx];
                        LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                        LocalColumnPermutationTracker[switchy] = hold;
                        hold = LocalBand1[switchx];
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalColumnPermutationTracker[switchx];
                        LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                        LocalColumnPermutationTracker[switchy] = hold;
                        hold = LocalBand1[switchx];
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalColumnPermutationTracker[switchx];
                        LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                        LocalColumnPermutationTracker[switchy] = hold;
                        hold = LocalBand1[switchx];
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 7;
                        switchy += 7;
                        hold = LocalBand1[switchx]; // row 2
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalBand1[switchx];
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalBand1[switchx];
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 7;
                        switchy += 7;
                        hold = LocalBand1[switchx]; // row 3
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalBand1[switchx];
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalBand1[switchx];
                        LocalBand1[switchx] = LocalBand1[switchy];
                        LocalBand1[switchy] = hold;
                     }

                     MinLexCandidateSw = true;
                     //Select Case FirstNonZeroDigitPositionInRow(2)
                     //Case < 1 // If -1 or 0, do nothing.
                     if (FirstNonZeroDigitPositionInRow[2] < 1) // If -1 or 0, do nothing.
                     {
                     }
                     //Case 5
                     else if (FirstNonZeroDigitPositionInRow[2] == 5)
                     {
                        if (Row2MiniRowCount[0] > 0 || LocalBand1[12] > 0 || LocalBand1[13] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 4
                     else if (FirstNonZeroDigitPositionInRow[2] == 4)
                     {
                        if (Row2MiniRowCount[0] > 0 || LocalBand1[12] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case > 5
                     else if (FirstNonZeroDigitPositionInRow[2] > 5)
                     {
                        if (Row2MiniRowCount[0] > 0 || Row2MiniRowCount[1] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 3
                     else if (FirstNonZeroDigitPositionInRow[2] == 3)
                     {
                        if (Row2MiniRowCount[0] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 2
                     else if (FirstNonZeroDigitPositionInRow[2] == 2)
                     {
                        if (LocalBand1[9] > 0 || LocalBand1[10] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 1
                     else if (FirstNonZeroDigitPositionInRow[2] == 1)
                     {
                        if (LocalBand1[9] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }

                     if (MinLexCandidateSw)
                     {
                        ColumnPermutationCode = Band1MiniRowColumnPermutationCode[0] * 16 + Band1MiniRowColumnPermutationCode[1] * 4 + Band1MiniRowColumnPermutationCode[2];
                        int tempVar2 = ColumnPermutations[ ColumnPermutationCode];
                        for (int columnpermutationix = 0; columnpermutationix <= tempVar2; columnpermutationix++)
                        {
                           if (columnpermutationix > 0) // Band1 (3 row) column permutation.
                           {
                              FirstColumnPermutationTrackerIsIdentitySw = false;
                              switchx = PermutationColumnX[ColumnPermutationCode][columnpermutationix];
                              switchy = PermutationColumnY[ColumnPermutationCode][columnpermutationix];
                              hold = LocalColumnPermutationTracker[switchx];
                              LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                              LocalColumnPermutationTracker[switchy] = hold;
                              hold = LocalBand1[switchx]; // row 1
                              LocalBand1[switchx] = LocalBand1[switchy];
                              LocalBand1[switchy] = hold;
                              switchx += 9; // row2
                              switchy += 9;
                              hold = LocalBand1[switchx];
                              LocalBand1[switchx] = LocalBand1[switchy];
                              LocalBand1[switchy] = hold;
                              switchx += 9; // row3
                              switchy += 9;
                              hold = LocalBand1[switchx];
                              LocalBand1[switchx] = LocalBand1[switchy];
                              LocalBand1[switchy] = hold;
                           }
                           else
                           {
                              for (i = 9; i <= 17; i++)
                              {
                                 if (LocalBand1[i] > 0)
                                 {
                                    Row2TestFirstNonZeroDigitPositionInRow = i - 9; // Note: 0-8 notation.
                                    if (FirstNonZeroDigitPositionInRow[2] < Row2TestFirstNonZeroDigitPositionInRow)
                                    {
                                       FirstNonZeroDigitPositionInRow[2] = Row2TestFirstNonZeroDigitPositionInRow;
                                    }
                                    break;
                                 }
                              }
                           }

                           if (FirstNonZeroDigitPositionInRow[2] == Row2TestFirstNonZeroDigitPositionInRow) // Check if can bypass relable check is no repeating digit in two rows. ??????
                           {
                              //Array.Clear(DigitAlreadyHitSw, 0, 10)
                              for (z = 0; z <= 9; z++)
                              {
                                 DigitAlreadyHitSw[z] = false;
                              }
                              //Array.Clear(TestRowsRelabeled, 0,  FirstNonZeroPositionInRow + 1)
                              for (z = 0; z <= FirstNonZeroPositionInRow1; z++)
                              {
                                 TestRowsRelabeled[z] = 0;
                              }
                              //Array.ConstrainedCopy(DigitsRelabelWrkInit, 0, DigitsRelabelWrk, 0, 10) ' Initialize DigitsRelabelWrk 1 to 9 to 10's (the zero element = 0).
                              for (z = 0; z <= 9; z++)
                              {
                                 DigitsRelabelWrk[z] = DigitsRelabelWrkInit[z];
                              }
                              RelabelLastDigit = 0;
                              for (i = FirstNonZeroPositionInRow1; i <= 26; i++) // Build DigitsRelabelWrk and TestGridRelabeled for Band1
                              {
                                 if (LocalBand1[i] > 0 && !DigitAlreadyHitSw[LocalBand1[i]])
                                 {
                                    DigitAlreadyHitSw[LocalBand1[i]] = true;
                                    RelabelLastDigit += 1;
                                    DigitsRelabelWrk[LocalBand1[i]] = RelabelLastDigit;
                                 }
                                 TestRowsRelabeled[i] = DigitsRelabelWrk[LocalBand1[i]];
                              }

                              for (i = FirstNonZeroPositionInRow1; i <= 26; i++)
                              {
                                 if (TestRowsRelabeled[i] > MinLexRows[i]) // Check if row2 is a candidate.
                                 {
                                    MinLexCandidateSw = false;
                                    break;
                                 }
                                 else if (TestRowsRelabeled[i] < MinLexRows[i])
                                 {
                                    break;
                                 }
                              }
                              if (MinLexCandidateSw)
                              {
                                 if (i < 27)
                                 {
                                    ArrayConstrainedCopy(TestRowsRelabeled, 0, MinLexRows, 0, 27);
                                    MinLexRowHitSw = false;
                                    ResetMinLexRowSw = true;
                                    ColumnPermutationTrackerIx = 0;
                                    ColumnPermutationTrackerIsIdentitySw[0] = FirstColumnPermutationTrackerIsIdentitySw;
                                    //Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, 0, 9)
                                    for (z = 0; z <= 8; z++)
                                    {
                                       ColumnPermutationTracker[z] = LocalColumnPermutationTracker[z];
                                    }
                                    //Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, 0, 10)
                                    zstart = ColumnPermutationTrackerIx * 10;
                                    for (z = 0; z <= 9; z++)
                                    {
                                       HoldDigitsRelabelWrk[z + zstart] = DigitsRelabelWrk[z];
                                    }
                                 }
                                 else
                                 {
                                    MinLexRowHitSw = true;
                                    ColumnPermutationTrackerIx += 1;
                                    if (FoundColumnPermutationTrackerCountMax < ColumnPermutationTrackerIx + 1)
                                    {
                                       FoundColumnPermutationTrackerCountMax = ColumnPermutationTrackerIx + 1;
                                    }
                                       ColumnPermutationTrackerIsIdentitySw[ColumnPermutationTrackerIx] = FirstColumnPermutationTrackerIsIdentitySw;
                                       //Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, ColumnPermutationTrackerIx * 9, 9)
                                       zstart = ColumnPermutationTrackerIx * 9;
                                       for (z = 0; z <= 8; z++)
                                       {
                                          ColumnPermutationTracker[z + zstart] = LocalColumnPermutationTracker[z];
                                       }
                                       //Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, ColumnPermutationTrackerIx * 10, 10)
                                       zstart = ColumnPermutationTrackerIx * 10;
                                       for (z = 0; z <= 9; z++)
                                       {
                                          HoldDigitsRelabelWrk[z + zstart] = DigitsRelabelWrk[z];
                                       }
                                 }
                              } // If MinLexCandidateSw
                           } // If FirstNonZeroDigitPositionInRow(2) = Row2TestFirstNonZeroDigitPositionInRow
                           MinLexCandidateSw = true;
                        }
                     } // If MinLexCandidateSw
                  }
               } // If MinLexCandidateSw

               if (ResetMinLexRowSw)
               {
                  Step2Row1CandidateIx = 0;
                  Step2Row1Candidate[0] = Row1Candidate;
                  Step2Row2Candidate[0] = Row2Candidate;
                  Step2ColumnPermutationTrackerStartIx[0] = -1;
                  Step2ColumnPermutationTrackerCount[0] = ColumnPermutationTrackerIx + 1;
                  ArrayConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, 0, 81);
                  //Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, 0, 10)
                  for (z = 0; z <= 9; z++)
                  {
                     HoldDigitAlreadyHitSw[z] = DigitAlreadyHitSw[z];
                  }
                  HoldRelabelLastDigit[0] = RelabelLastDigit;
               }
               else if (MinLexRowHitSw)
               {
                  Step2Row1CandidateIx += 1;
                  Step2Row1Candidate[Step2Row1CandidateIx] = Row1Candidate;
                  Step2Row2Candidate[Step2Row1CandidateIx] = Row2Candidate;
                  Step2ColumnPermutationTrackerStartIx[Step2Row1CandidateIx] = CandidateColumnPermutationTrackerStartIx;
                  Step2ColumnPermutationTrackerCount[Step2Row1CandidateIx] = ColumnPermutationTrackerIx - CandidateColumnPermutationTrackerStartIx;
                  ArrayConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, Step2Row1CandidateIx * 81, 81);
                  //Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, Step2Row1CandidateIx * 10, 10)
                  zstart = Step2Row1CandidateIx * 10;
                  for (z = 0; z <= 9; z++)
                  {
                     HoldDigitAlreadyHitSw[z + zstart] = DigitAlreadyHitSw[z];
                  }
                  HoldRelabelLastDigit[Step2Row1CandidateIx] = RelabelLastDigit;
               }

            }
            ArrayConstrainedCopy(MinLexRows, 0, MinLexGridLocal, 0, 27); // For cases of non-zero row1, Copy MinLexRows to MinLexGridLocal.
            // End MinLexBand1 (Row1 not empty case)
         }
         else // Start MinLexRows2and3 (Row1 empty case)
         {
            FirstNonZeroDigitPositionInRow[3] = -1;
            Row2StackPermutationCode = FirstNonZeroRowStackPermutationCode[Row2CalcMiniRowCountCodeMinimum];
            //Array.ConstrainedCopy(MinLexGridLocal, 9, MinLexRows, 0, 18) ' Copy first non-zero row to MinLexRows row 1; set second row to all 10s.
            for (z = 0; z <= 17; z++)
            {
               MinLexRows[z] = MinLexGridLocal[z + 9];
            }
            ColumnPermutationTrackerIx = -1;
            for (int candidate = 0; candidate <= Band1CandidateIx; candidate++)
            {
               Row1Candidate17 = Band1CandidateRow1[candidate];
               if (Row1Candidate17 < 9)
               {
                  Row2Candidate17 = Band1CandidateRow2[candidate];
                  Row3Candidate17 = Band1CandidateRow3[candidate];
                  Row1Candidate = Row1Candidate17 + 1;
                  Row2Candidate = Band1CandidateRow2[candidate] + 1;
                  Row3Candidate = Band1CandidateRow3[candidate] + 1;
                  ArrayConstrainedCopy(InputGrid, 0, HoldPuzzle, 0, 81); // Copy the direct input puzzle to HoldPuzzle.
               }
               else
               {
                  Row2Candidate17 = Band1CandidateRow2[candidate];
                  Row3Candidate17 = Band1CandidateRow3[candidate];
                  Row1Candidate = Row1Candidate17 - 8;
                  Row2Candidate = Band1CandidateRow2[candidate] - 8;
                  Row3Candidate = Band1CandidateRow3[candidate] - 8;
                  ArrayConstrainedCopy(InputGrid, 81, HoldPuzzle, 0, 81); // Copy the transposed input puzzle to HoldPuzzle.
               }

               // MinLex Rows 2 & 3 for Row2Candidate.

               MinLexRowHitSw = false;
               ResetMinLexRowSw = false;
               CandidateColumnPermutationTrackerStartIx = ColumnPermutationTrackerIx;
               row1start = (Row1Candidate - 1) * 9;
               row2start = (Row2Candidate - 1) * 9;
               row3start = (Row3Candidate - 1) * 9;

               StackPermutationTrackerLocal[0] = MiniRowOrderTracker[Row2Candidate17][0];
               StackPermutationTrackerLocal[1] = MiniRowOrderTracker[Row2Candidate17][1];
               StackPermutationTrackerLocal[2] = MiniRowOrderTracker[Row2Candidate17][2];
               Row2MiniRowCount[0] = JustifiedMiniRowCount[Row2Candidate17][0];
               Row2MiniRowCount[1] = JustifiedMiniRowCount[Row2Candidate17][1];
               Row2MiniRowCount[2] = JustifiedMiniRowCount[Row2Candidate17][2];
               Row3MiniRowCount[0] = OriginalMiniRowCount[Row3Candidate17][StackPermutationTrackerLocal[0]];
               Row3MiniRowCount[1] = OriginalMiniRowCount[Row3Candidate17][StackPermutationTrackerLocal[1]];
               Row3MiniRowCount[2] = OriginalMiniRowCount[Row3Candidate17][StackPermutationTrackerLocal[2]];

               Row3StackPermutationCode = Row2StackPermutationCode;
               if (Row2CalcMiniRowCountCodeMinimum < 4 || Row2StackPermutationCode > 1)
               {
                  if (Row3MiniRowCount[0] > Row3MiniRowCount[1]) // If Row3 minirow1 count is greater than minirow2, switch minirow1 and minirow2.
                  {
                     hold = Row3MiniRowCount[0];
                     Row3MiniRowCount[0] = Row3MiniRowCount[1];
                     Row3MiniRowCount[1] = hold;
                     hold = StackPermutationTrackerLocal[0];
                     StackPermutationTrackerLocal[0] = StackPermutationTrackerLocal[1];
                     StackPermutationTrackerLocal[1] = hold;
                  }
                  else if (Row2CalcMiniRowCountCodeMinimum < 4 && Row3MiniRowCount[1] > 0 && Row3MiniRowCount[0] == Row3MiniRowCount[1])
                  {
                     Row3StackPermutationCode = 2;
                  }
               }

               MinLexCandidateSw = true;
               if (Row2StackPermutationCode == 0)
               {
                  //Select Case FirstNonZeroDigitPositionInRow(3)
                  //Case < 1 // If -1 or 0, do nothing.
                  if (FirstNonZeroDigitPositionInRow[3] < 1) // If -1 or 0, do nothing.
                  {
                  }
                  //Case 5
                  else if (FirstNonZeroDigitPositionInRow[3] == 5)
                  {
                     if (Row3MiniRowCount[0] > 0 || Row3MiniRowCount[1] > 1)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 4
                  else if (FirstNonZeroDigitPositionInRow[3] == 4)
                  {
                     if (Row3MiniRowCount[0] > 0 || Row3MiniRowCount[1] == 3)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case > 5
                  else if (FirstNonZeroDigitPositionInRow[3] > 5)
                  {
                     if (Row3MiniRowCount[0] > 0 || Row3MiniRowCount[1] > 0)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 3
                  else if (FirstNonZeroDigitPositionInRow[3] == 3)
                  {
                     if (Row3MiniRowCount[0] > 0)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 2
                  else if (FirstNonZeroDigitPositionInRow[3] == 2)
                  {
                     if (Row3MiniRowCount[0] > 1)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
                  //Case 1
                  else if (FirstNonZeroDigitPositionInRow[3] == 1)
                  {
                     if (Row3MiniRowCount[0] == 3)
                     {
                        MinLexCandidateSw = false;
                     }
                  }
               }

               if (MinLexCandidateSw)
               {
                  StackPermutationTrackerCode = StackPermutationTrackerLocal[0] * 100 + StackPermutationTrackerLocal[1] * 10 + StackPermutationTrackerLocal[2];
                  switch (StackPermutationTrackerCode)
                  {
                     case 12:
                     break;
                     case 21:
                        SwitchStacks12(HoldPuzzle);
                        break;
                     case 102:
                        SwitchStacks01(HoldPuzzle);
                        break;
                     case 120:
                        Switch3Stacks120(HoldPuzzle);
                        break;
                     case 201:
                        Switch3Stacks201(HoldPuzzle);
                        break;
                     case 210:
                        SwitchStacks02(HoldPuzzle);
                        break;
                  }

                  // Positionally (not considering digit values) "right justify" second and third row non-zero digits within MiniRows.
                  // Also set the ColumnPermutationCode for Band1.
                  // For each MiniRow (0 to 2 notation).
                  Band1MiniRowColumnPermutationCode[0] = 0;
                  Band1MiniRowColumnPermutationCode[1] = 0;
                  Band1MiniRowColumnPermutationCode[2] = 0;
                  switch (Row2MiniRowCount[0]) // First MiniRow
                  {
                     case 0: // Row2
                        switch (Row3MiniRowCount[0])
                        {
                           case 0:
                           break;
                           case 1:
                              if (HoldPuzzle[row3start] > 0)
                              {
                                 SwitchColumns02(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row3start + 1] > 0)
                              {
                                 SwitchColumns12(HoldPuzzle);
                              }
                              break;
                           case 2:
                              if (HoldPuzzle[row3start + 2] == 0)
                              {
                                 SwitchColumns02(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row3start + 1] == 0)
                              {
                                 SwitchColumns01(HoldPuzzle);
                              }
                              Band1MiniRowColumnPermutationCode[0] += 1;
                              break;
                           case 3:
                              Band1MiniRowColumnPermutationCode[0] += 3;
                              break;
                        }
                        break;
                     case 1: // Row 2
                        if (HoldPuzzle[row2start] > 0)
                        {
                           SwitchColumns02(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row2start + 1] > 0)
                        {
                           SwitchColumns12(HoldPuzzle);
                        }
                        if (HoldPuzzle[row3start] > 0)
                        {
                           if (HoldPuzzle[row3start + 1] == 0)
                           {
                              SwitchColumns01(HoldPuzzle);
                           }
                           else
                           {
                              Band1MiniRowColumnPermutationCode[0] += 2;
                           }
                        }
                        break;
                     case 2: // Row 2
                        if (HoldPuzzle[row2start + 2] == 0)
                        {
                           SwitchColumns02(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row2start + 1] == 0)
                        {
                           SwitchColumns01(HoldPuzzle);
                        }
                        Band1MiniRowColumnPermutationCode[0] += 1;
                        break;
                     case 3: // Row 2
                        Band1MiniRowColumnPermutationCode[0] += 3;
                        break;
                  }

                  switch (Row2MiniRowCount[1]) // Second MiniRow
                  {
                     case 0: // Row2
                        switch (Row3MiniRowCount[1])
                        {
                           case 1:
                              if (HoldPuzzle[row3start + 3] > 0)
                              {
                                 SwitchColumns35(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row3start + 4] > 0)
                              {
                                 SwitchColumns45(HoldPuzzle);
                              }
                              break;
                           case 2:
                              if (HoldPuzzle[row3start + 5] == 0)
                              {
                                 SwitchColumns35(HoldPuzzle);
                              }
                              else if (HoldPuzzle[row3start + 4] == 0)
                              {
                                 SwitchColumns34(HoldPuzzle);
                              }
                              Band1MiniRowColumnPermutationCode[1] += 1;
                              break;
                           case 3:
                              Band1MiniRowColumnPermutationCode[1] += 3;
                              break;
                        }
                        break;
                     case 1: // Row2
                        if (HoldPuzzle[row2start + 3] > 0)
                        {
                           SwitchColumns35(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row2start + 4] > 0)
                        {
                           SwitchColumns45(HoldPuzzle);
                        }
                        if (HoldPuzzle[row3start + 3] > 0)
                        {
                           if (HoldPuzzle[row3start + 4] == 0)
                           {
                              SwitchColumns34(HoldPuzzle);
                           }
                           else
                           {
                              Band1MiniRowColumnPermutationCode[1] += 2;
                           }
                        }
                        break;
                     case 2: // Row2
                        if (HoldPuzzle[row2start + 5] == 0)
                        {
                           SwitchColumns35(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row2start + 4] == 0)
                        {
                           SwitchColumns34(HoldPuzzle);
                        }
                        Band1MiniRowColumnPermutationCode[1] += 1;
                        break;
                     case 3: // Row2
                        Band1MiniRowColumnPermutationCode[1] += 3;
                        break;
                  }
                  switch (Row2MiniRowCount[2]) // Third MiniRow
                  {
                     case 1: // Row2
                        if (HoldPuzzle[row2start + 6] > 0)
                        {
                           SwitchColumns68(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row2start + 7] > 0)
                        {
                           SwitchColumns78(HoldPuzzle);
                        }
                        if (HoldPuzzle[row3start + 6] > 0)
                        {
                           if (HoldPuzzle[row3start + 7] == 0)
                           {
                              SwitchColumns67(HoldPuzzle);
                           }
                           else
                           {
                              Band1MiniRowColumnPermutationCode[2] += 2;
                           }
                        }
                        break;
                     case 2: // Row2
                        if (HoldPuzzle[row2start + 8] == 0)
                        {
                           SwitchColumns68(HoldPuzzle);
                        }
                        else if (HoldPuzzle[row2start + 7] == 0)
                        {
                           SwitchColumns67(HoldPuzzle);
                        }
                        Band1MiniRowColumnPermutationCode[2] += 1;
                        break;
                     case 3: // Row2
                        Band1MiniRowColumnPermutationCode[2] += 3;
                        break;
                  }

                  //Array.ConstrainedCopy(HoldPuzzle, (Row2Candidate - 1) * 9, LocalRows2and3, 0, 9)
                  zstart = (Row2Candidate - 1) * 9;
                  for (z = 0; z <= 8; z++)
                  {
                     LocalRows2and3[z] = HoldPuzzle[z + zstart];
                  }
                  //Array.ConstrainedCopy(HoldPuzzle, (Row3Candidate - 1) * 9, LocalRows2and3, 9, 9)
                  zstart = (Row3Candidate - 1) * 9;
                  for (z = 0; z <= 8; z++)
                  {
                     LocalRows2and3[z + 9] = HoldPuzzle[z + zstart];
                  }
                  //Array.ConstrainedCopy(ColumnTrackerInit, 0, LocalColumnPermutationTracker, 0, 9)
                  for (z = 0; z <= 8; z++)
                  {
                     LocalColumnPermutationTracker[z] = ColumnTrackerInit[z];
                  }
                  FirstColumnPermutationTrackerIsIdentitySw = true;
                  int tempVar3 = StackPermutations[ Row3StackPermutationCode];
                  for (int stackpermutationix = 0; stackpermutationix <= tempVar3; stackpermutationix++)
                  {
                     if (stackpermutationix > 0)
                     {
                        FirstColumnPermutationTrackerIsIdentitySw = false;
                        stackx = PermutationStackX[Row3StackPermutationCode][stackpermutationix]; // Band1 (2 row) stack switch.
                        stacky = PermutationStackY[Row3StackPermutationCode][stackpermutationix];
                        switchx = 3 * stackx;
                        switchy = 3 * stacky;
                        hold = Band1MiniRowColumnPermutationCode[stackx];
                        Band1MiniRowColumnPermutationCode[stackx] = Band1MiniRowColumnPermutationCode[stacky];
                        Band1MiniRowColumnPermutationCode[stacky] = hold;
                        hold = Row3MiniRowCount[stackx];
                        Row3MiniRowCount[stackx] = Row3MiniRowCount[stacky];
                        Row3MiniRowCount[stacky] = hold;
                        hold = LocalColumnPermutationTracker[switchx];
                        LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                        LocalColumnPermutationTracker[switchy] = hold;
                        hold = LocalRows2and3[switchx];
                        LocalRows2and3[switchx] = LocalRows2and3[switchy];
                        LocalRows2and3[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalColumnPermutationTracker[switchx];
                        LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                        LocalColumnPermutationTracker[switchy] = hold;
                        hold = LocalRows2and3[switchx];
                        LocalRows2and3[switchx] = LocalRows2and3[switchy];
                        LocalRows2and3[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalColumnPermutationTracker[switchx];
                        LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                        LocalColumnPermutationTracker[switchy] = hold;
                        hold = LocalRows2and3[switchx];
                        LocalRows2and3[switchx] = LocalRows2and3[switchy];
                        LocalRows2and3[switchy] = hold;
                        switchx += 7;
                        switchy += 7;
                        hold = LocalRows2and3[switchx];
                        LocalRows2and3[switchx] = LocalRows2and3[switchy];
                        LocalRows2and3[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalRows2and3[switchx];
                        LocalRows2and3[switchx] = LocalRows2and3[switchy];
                        LocalRows2and3[switchy] = hold;
                        switchx += 1;
                        switchy += 1;
                        hold = LocalRows2and3[switchx];
                        LocalRows2and3[switchx] = LocalRows2and3[switchy];
                        LocalRows2and3[switchy] = hold;
                     }

                     MinLexCandidateSw = true;
                     //Select Case FirstNonZeroDigitPositionInRow(3)
                     //Case < 1 // If -1 or 0, do nothing.
                     if (FirstNonZeroDigitPositionInRow[3] < 1) // If -1 or 0, do nothing.
                     {
                     }
                     //Case 5
                     else if (FirstNonZeroDigitPositionInRow[3] == 5)
                     {
                        if (Row3MiniRowCount[0] > 0 || LocalRows2and3[12] > 0 || LocalRows2and3[13] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 4
                     else if (FirstNonZeroDigitPositionInRow[3] == 4)
                     {
                        if (Row3MiniRowCount[0] > 0 || LocalRows2and3[12] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case > 5
                     else if (FirstNonZeroDigitPositionInRow[3] > 5)
                     {
                        if (Row3MiniRowCount[0] > 0 || Row3MiniRowCount[1] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 3
                     else if (FirstNonZeroDigitPositionInRow[3] == 3)
                     {
                        if (Row3MiniRowCount[0] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 2
                     else if (FirstNonZeroDigitPositionInRow[3] == 2)
                     {
                        if (LocalRows2and3[9] > 0 || LocalRows2and3[10] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }
                     //Case 1
                     else if (FirstNonZeroDigitPositionInRow[3] == 1)
                     {
                        if (LocalRows2and3[9] > 0)
                        {
                           MinLexCandidateSw = false;
                        }
                     }

                     if (MinLexCandidateSw)
                     {
                        ColumnPermutationCode = Band1MiniRowColumnPermutationCode[0] * 16 + Band1MiniRowColumnPermutationCode[1] * 4 + Band1MiniRowColumnPermutationCode[2];
                        int tempVar4 = ColumnPermutations[ ColumnPermutationCode];
                        for (int columnpermutationix = 0; columnpermutationix <= tempVar4; columnpermutationix++)
                        {
                           if (columnpermutationix > 0) // Band1 (2 row) column permutation.
                           {
                              FirstColumnPermutationTrackerIsIdentitySw = false;
                              switchx = PermutationColumnX[ColumnPermutationCode][columnpermutationix];
                              switchy = PermutationColumnY[ColumnPermutationCode][columnpermutationix];
                              hold = LocalColumnPermutationTracker[switchx];
                              LocalColumnPermutationTracker[switchx] = LocalColumnPermutationTracker[switchy];
                              LocalColumnPermutationTracker[switchy] = hold;
                              hold = LocalRows2and3[switchx]; // row 2
                              LocalRows2and3[switchx] = LocalRows2and3[switchy];
                              LocalRows2and3[switchy] = hold;
                              switchx += 9;
                              switchy += 9;
                              hold = LocalRows2and3[switchx]; // row 3
                              LocalRows2and3[switchx] = LocalRows2and3[switchy];
                              LocalRows2and3[switchy] = hold;
                           }
                           else
                           {
                              for (i = 9; i <= 17; i++)
                              {
                                 if (LocalRows2and3[i] > 0)
                                 {
                                    Row3TestFirstNonZeroDigitPositionInRow = i - 9; // Note: 0-8 notation.
                                    if (FirstNonZeroDigitPositionInRow[3] < Row3TestFirstNonZeroDigitPositionInRow)
                                    {
                                       FirstNonZeroDigitPositionInRow[3] = Row3TestFirstNonZeroDigitPositionInRow;
                                    }
                                    break;
                                 }
                              }
                           }

                           if (FirstNonZeroDigitPositionInRow[3] == Row3TestFirstNonZeroDigitPositionInRow)
                           {
                              //Array.Clear(DigitAlreadyHitSw, 0, 10)
                              for (z = 0; z <= 9; z++)
                              {
                                 DigitAlreadyHitSw[z] = false;
                              }
                              //Array.Clear(TestRowsRelabeled, 0, FirstNonZeroPositionInRow + 1)
                              for (z = 0; z <= FirstNonZeroPositionInRow2; z++)
                              {
                                 TestRowsRelabeled[z] = 0;
                              }
                              //Array.ConstrainedCopy(DigitsRelabelWrkInit, 0, DigitsRelabelWrk, 0, 10) ' Initialize DigitsRelabelWrk 1 to 9 to 10's (the zero element = 0).
                              for (z = 0; z <= 9; z++)
                              {
                                 DigitsRelabelWrk[z] = DigitsRelabelWrkInit[z];
                              }
                              RelabelLastDigit = 0;
                              for (i = FirstNonZeroPositionInRow2; i <= 17; i++) // Build DigitsRelabelWrk and TestGridRelabeled for
                              {
                                 if (LocalRows2and3[i] > 0 && !DigitAlreadyHitSw[LocalRows2and3[i]])
                                 {
                                    DigitAlreadyHitSw[LocalRows2and3[i]] = true;
                                    RelabelLastDigit += 1;
                                    DigitsRelabelWrk[LocalRows2and3[i]] = RelabelLastDigit;
                                 }
                                 TestRowsRelabeled[i] = DigitsRelabelWrk[LocalRows2and3[i]];
                              }
                              for (i = FirstNonZeroPositionInRow2; i <= 17; i++)
                              {
                                 if (TestRowsRelabeled[i] > MinLexRows[i]) // Check if row2 is a candidate.
                                 {
                                    MinLexCandidateSw = false;
                                    break;
                                 }
                                 else if (TestRowsRelabeled[i] < MinLexRows[i])
                                 {
                                    break;
                                 }
                              }
                              if (MinLexCandidateSw)
                              {
                                 if (i < 18)
                                 {
                                    //Array.ConstrainedCopy(TestRowsRelabeled, 0, MinLexRows, 0, 18)
                                    for (z = 0; z <= 17; z++)
                                    {
                                       MinLexRows[z] = TestRowsRelabeled[z];
                                    }
                                    MinLexRowHitSw = false;
                                    ResetMinLexRowSw = true;
                                    ColumnPermutationTrackerIx = 0;
                                    ColumnPermutationTrackerIsIdentitySw[0] = FirstColumnPermutationTrackerIsIdentitySw;
                                    //Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, 0, 9)
                                    for (z = 0; z <= 8; z++)
                                    {
                                       ColumnPermutationTracker[z] = LocalColumnPermutationTracker[z];
                                    }
                                    //Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, 0, 10)
                                    for (z = 0; z <= 9; z++)
                                    {
                                       HoldDigitsRelabelWrk[z] = DigitsRelabelWrk[z];
                                    }
                                 }
                                 else
                                 {
                                    MinLexRowHitSw = true;
                                    ColumnPermutationTrackerIx += 1;
                                    if (FoundColumnPermutationTrackerCountMax < ColumnPermutationTrackerIx)
                                    {
                                       FoundColumnPermutationTrackerCountMax = ColumnPermutationTrackerIx + 1;
                                    }
                                       ColumnPermutationTrackerIsIdentitySw[ColumnPermutationTrackerIx] = FirstColumnPermutationTrackerIsIdentitySw;
                                       //Array.ConstrainedCopy(LocalColumnPermutationTracker, 0, ColumnPermutationTracker, ColumnPermutationTrackerIx * 9, 9)
                                       zstart = ColumnPermutationTrackerIx * 9;
                                       for (z = 0; z <= 8; z++)
                                       {
                                          ColumnPermutationTracker[z + zstart] = LocalColumnPermutationTracker[z];
                                       }
                                       //Array.ConstrainedCopy(DigitsRelabelWrk, 0, HoldDigitsRelabelWrk, 0, 10)
                                       zstart = ColumnPermutationTrackerIx * 10;
                                       for (z = 0; z <= 9; z++)
                                       {
                                          HoldDigitsRelabelWrk[z + zstart] = DigitsRelabelWrk[z];
                                       }
                                 }
                              } // If MinLexCandidateSw
                           } // If FirstNonZeroDigitPositionInRow(3) = Row3TestFirstNonZeroDigitPositionInRow
                           MinLexCandidateSw = true;
                        }
                     } // If MinLexCandidateSw
                  }
               } // If MinLexCandidateSw

               if (ResetMinLexRowSw)
               {
                  Step2Row1CandidateIx = 0;
                  Step2Row1Candidate[0] = Row1Candidate;
                  Step2Row2Candidate[0] = Row2Candidate;
                  Step2ColumnPermutationTrackerStartIx[0] = -1;
                  Step2ColumnPermutationTrackerCount[0] = ColumnPermutationTrackerIx + 1;
                  ArrayConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, 0, 81);
                  //Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, 0, 10)
                  for (z = 0; z <= 9; z++)
                  {
                     HoldDigitAlreadyHitSw[z] = DigitAlreadyHitSw[z];
                  }
                  HoldRelabelLastDigit[0] = RelabelLastDigit;
               }
               else if (MinLexRowHitSw)
               {
                  Step2Row1CandidateIx += 1;
                  Step2Row1Candidate[Step2Row1CandidateIx] = Row1Candidate;
                  Step2Row2Candidate[Step2Row1CandidateIx] = Row2Candidate;
                  Step2ColumnPermutationTrackerStartIx[Step2Row1CandidateIx] = CandidateColumnPermutationTrackerStartIx;
                  Step2ColumnPermutationTrackerCount[Step2Row1CandidateIx] = ColumnPermutationTrackerIx - CandidateColumnPermutationTrackerStartIx;
                  ArrayConstrainedCopy(HoldPuzzle, 0, HoldBand1CandidateJustifiedPuzzles, Step2Row1CandidateIx * 81, 81);
                  //Array.ConstrainedCopy(DigitAlreadyHitSw, 0, HoldDigitAlreadyHitSw, Step2Row1CandidateIx * 10, 10)
                  zstart = Step2Row1CandidateIx * 10;
                  for (z = 0; z <= 9; z++)
                  {
                     HoldDigitAlreadyHitSw[z + zstart] = DigitAlreadyHitSw[z];
                  }
                  HoldRelabelLastDigit[Step2Row1CandidateIx] = RelabelLastDigit;
               }
            }
            //Array.ConstrainedCopy(MinLexRows, 0, MinLexGridLocal, 9, 18) ' For cases of all zeros row1, zero out the first row of MinLexGridLocal and copy MinLexRows to the second and third row of MinLexGridLocal.
            for (z = 0; z <= 8; z++)
            {
               MinLexGridLocal[z] = 0;
            }
            for (z = 0; z <= 17; z++)
            {
               MinLexGridLocal[z + 9] = MinLexRows[z];
            }
            // End MinLexRows2and3 (Row1 empty case)
         } // If CalcMiniRowCountCodeMinimum > 0 Then

         CheckThisPassSw = true;
         FirstNonZeroRowCandidateIx = 0;
         ArrayConstrainedCopy(MinLexGridLocal, 0, TestGridRelabeled, 0, 27); // Copy MinLexGridLocal Band1 to TestGridRelebeled.
         while (FirstNonZeroRowCandidateIx <= Step2Row1CandidateIx)
         {
            Row1Candidate = Step2Row1Candidate[FirstNonZeroRowCandidateIx];
            Row2Candidate = Step2Row2Candidate[FirstNonZeroRowCandidateIx];
            ArrayConstrainedCopy(HoldBand1CandidateJustifiedPuzzles, FirstNonZeroRowCandidateIx * 81, TestBand1, 0, 81);
            //Array.ConstrainedCopy(HoldDigitAlreadyHitSw, FirstNonZeroRowCandidateIx * 10, Row3DigitAlreadyHitSw, 0, 10)
            zstart = FirstNonZeroRowCandidateIx * 10;
            for (z = 0; z <= 9; z++)
            {
               Row3DigitAlreadyHitSw[z] = HoldDigitAlreadyHitSw[z + zstart];
            }
            Row3RelabelLastDigit = HoldRelabelLastDigit[FirstNonZeroRowCandidateIx];

            // Move next row1, row2 and row3 candidates to rows 1, 2 and 3 respectively.
            ArrayConstrainedCopy(TestBand1, 0, HoldGrid, 0, 81);
            switch (Row1Candidate)
            {
               case 1: // If next row1 candidate is row 1 and the row2 candidate is row 2 then do nothing.
                  if (Row2Candidate == 3) // If next row1 candidate is row 1 and the row2 candidate is row 3 then switch row 2 and row 3.
                  {
                     //Array.ConstrainedCopy(TestBand1, 9, HoldGrid, 18, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z + 9] = TestBand1[z];
                     }
                     //Array.ConstrainedCopy(TestBand1, 18, HoldGrid, 9, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 9];
                     }
                  }
                  break;
               case 2: // If next row1 candidate is row 2, ...
                  if (Row2Candidate == 1) // if the row2 candidate is row 1 then switch rows 1 and 2,
                  {
                     //Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 9, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z + 9] = TestBand1[z];
                     }
                     //Array.ConstrainedCopy(TestBand1, 9, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 9];
                     }
                  }
                  else // else the row2 candidate is row 3 so move rows 1, 2 and 3 to rows 3, 1 and 2.
                  {
                     //Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 18, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z + 18] = TestBand1[z];
                     }
                     //Array.ConstrainedCopy(TestBand1, 9, HoldGrid, 0, 18)
                     for (z = 0; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 9];
                     }
                  }
                  break;
               case 3: // If next row1 candidate is row 3, ...
                  if (Row2Candidate == 1) // if the row2 candidate is row 1 then move rows 1, 2 and 3 to rows 2, 3 and 1,
                  {
                     //    Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 9, 18)
                     for (z = 0; z <= 17; z++)
                     {
                        HoldGrid[z + 9] = TestBand1[z];
                     }
                     //Array.ConstrainedCopy(TestBand1, 18, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 18];
                     }
                  }
                  else // else the row2 candidate is row 2 so switch rows 1 and 3.
                  {
                     //Array.ConstrainedCopy(TestBand1, 0, HoldGrid, 18, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z + 18] = TestBand1[z];
                     }
                     //Array.ConstrainedCopy(TestBand1, 18, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 18];
                     }
                  }
                  break;
               case 4: // If next row1 candidate is row 4 move band 1 to band 2, then ...
                  ArrayConstrainedCopy(TestBand1, 0, HoldGrid, 27, 27);
                  if (Row2Candidate == 5) // if the row2 candidate is row 5 then move band 2 to band 1,
                  {
                     ArrayConstrainedCopy(TestBand1, 27, HoldGrid, 0, 27);
                  }
                  else // else the row2 candidate is row 6 so move rows 4, 5 and 6 to rows 1, 3, and 2.
                  {
                     //Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 27];
                     }
                     //Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 18];
                     }
                     //Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 9, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 36];
                     }
                  }
                  break;
               case 5: // If next row1 candidate is row 5 move band 1 to band 2, then ...
                  ArrayConstrainedCopy(TestBand1, 0, HoldGrid, 27, 27);
                  if (Row2Candidate == 4) // if the row2 candidate is row 4 then move rows 4, 5 and 6 to rows 2, 1 and 3,
                  {
                     //Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 9, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 18];
                     }
                     //Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 36];
                     }
                     //Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 27];
                     }
                  }
                  else // else the row2 candidate is row 6 so move rows 4, 5 and 6 to rows 3, 1 and 2.
                  {
                     //Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 9];
                     }
                     //Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 0, 18)
                     for (z = 0; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 36];
                     }
                  }
                  break;
               case 6: // If next row1 candidate is row 6 move band 1 to band 2, then ...
                  ArrayConstrainedCopy(TestBand1, 0, HoldGrid, 27, 27);
                  if (Row2Candidate == 4) // if the row2 candidate is row 4 then move rows 4, 5 and 6 to rows 2, 3 and 1,
                  {
                     //Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 9, 18)
                     for (z = 9; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 18];
                     }
                     //Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 45];
                     }
                  }
                  else // else the row2 candidate is row 5 so move rows 4, 5 and 6 to rows 3, 2 and 1.
                  {
                     //Array.ConstrainedCopy(TestBand1, 27, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 9];
                     }
                     //Array.ConstrainedCopy(TestBand1, 36, HoldGrid, 9, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 27];
                     }
                     //Array.ConstrainedCopy(TestBand1, 45, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 45];
                     }
                  }
                  break;
               case 7: // If next row1 candidate is row 7 move band 1 to band 3, then ...
                  ArrayConstrainedCopy(TestBand1, 0, HoldGrid, 54, 27);
                  if (Row2Candidate == 8) // if the row2 candidate is row 8 then move band 3 to band 1,
                  {
                     ArrayConstrainedCopy(TestBand1, 54, HoldGrid, 0, 27);
                  }
                  else // else the row2 candidate is row 9 so move rows 7, 8 and 9 to rows 1, 3, and 2.
                  {
                     //Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 54];
                     }
                     //Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 45];
                     }
                     //Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 9, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 63];
                     }
                  }
                  break;
               case 8: // If next row1 candidate is row 8 move band 1 to band 3, then ...
                  ArrayConstrainedCopy(TestBand1, 0, HoldGrid, 54, 27);
                  if (Row2Candidate == 7) // if the row2 candidate is row 7 then move rows 7, 8 and 9 to rows 2, 1 and 3,
                  {
                     //Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 9, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 45];
                     }
                     //Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 63];
                     }
                     //Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 54];
                     }
                  }
                  else // else the row2 candidate is row 9 so move rows 7, 8 and 9 to rows 3, 1 and 2.
                  {
                     //Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 36];
                     }
                     //Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 0, 18)
                     for (z = 0; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 63];
                     }
                  }
                  break;
               case 9: // If next row1 candidate is row 9 move band 1 to band 3, then ...
                  ArrayConstrainedCopy(TestBand1, 0, HoldGrid, 54, 27);
                  if (Row2Candidate == 7) // if the row2 candidate is row 7 then move rows 7, 8 and 9 to rows 2, 3 and 1,
                  {
                     //Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 9, 18)
                     for (z = 9; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 45];
                     }
                     //Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 72];
                     }
                  }
                  else // else the row2 candidate is row 8 so move rows 7, 8 and 9 to rows 3, 2 and 1.
                  {
                     //Array.ConstrainedCopy(TestBand1, 54, HoldGrid, 18, 9)
                     for (z = 18; z <= 26; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 36];
                     }
                     //Array.ConstrainedCopy(TestBand1, 63, HoldGrid, 9, 9)
                     for (z = 9; z <= 17; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 54];
                     }
                     //Array.ConstrainedCopy(TestBand1, 72, HoldGrid, 0, 9)
                     for (z = 0; z <= 8; z++)
                     {
                        HoldGrid[z] = TestBand1[z + 72];
                     }
                  }
                  break;
            }
            ColumnPermutationTrackerCount = Step2ColumnPermutationTrackerCount[FirstNonZeroRowCandidateIx];
            FirstColumnPermutationTrackerIsIdentitySw = ColumnPermutationTrackerIsIdentitySw[Step2ColumnPermutationTrackerStartIx[FirstNonZeroRowCandidateIx] + 1];
            for (int trackercolumnpermutationix = 0; trackercolumnpermutationix < ColumnPermutationTrackerCount; trackercolumnpermutationix++)
            {
               //Array.ConstrainedCopy(ColumnPermutationTracker, Step2ColumnPermutationTrackerStartIx(FirstNonZeroRowCandidateIx) * 9 + 9 * trackercolumnpermutationix, ColumnPermutationTracker, 0, 9)
               zstart = (Step2ColumnPermutationTrackerStartIx[FirstNonZeroRowCandidateIx] + 1) * 9 + 9 * trackercolumnpermutationix;
               for (z = 0; z <= 8; z++)
               {
                  LocalColumnPermutationTracker[z] = ColumnPermutationTracker[z + zstart];
               }
               //Array.ConstrainedCopy(HoldDigitsRelabelWrk, Step2DigitsRelabelWrkStartIx(FirstNonZeroRowCandidateIx) * 10 + 10 * trackercolumnpermutationix, Row3DigitsRelabelWrk, 0, 10)
               zstart = (Step2ColumnPermutationTrackerStartIx[FirstNonZeroRowCandidateIx] + 1) * 10 + 10 * trackercolumnpermutationix;
               for (z = 0; z <= 9; z++)
               {
                  Row3DigitsRelabelWrk[z] = HoldDigitsRelabelWrk[z + zstart];
               }
               ArrayConstrainedCopy(HoldGrid, 0, TestBand1, 0, 81);
               if (trackercolumnpermutationix > 0 || !FirstColumnPermutationTrackerIsIdentitySw)
               {
                  if (LocalColumnPermutationTracker[0] != 0) // Column 1
                  {
                     TestBand1[0] = HoldGrid[LocalColumnPermutationTracker[0]];
                     TestBand1[9] = HoldGrid[LocalColumnPermutationTracker[0] + 9];
                     TestBand1[18] = HoldGrid[LocalColumnPermutationTracker[0] + 18];
                     TestBand1[27] = HoldGrid[LocalColumnPermutationTracker[0] + 27];
                     TestBand1[36] = HoldGrid[LocalColumnPermutationTracker[0] + 36];
                     TestBand1[45] = HoldGrid[LocalColumnPermutationTracker[0] + 45];
                     TestBand1[54] = HoldGrid[LocalColumnPermutationTracker[0] + 54];
                     TestBand1[63] = HoldGrid[LocalColumnPermutationTracker[0] + 63];
                     TestBand1[72] = HoldGrid[LocalColumnPermutationTracker[0] + 72];
                  }
                  if (LocalColumnPermutationTracker[1] != 1) // Column 2
                  {
                     TestBand1[1] = HoldGrid[LocalColumnPermutationTracker[1]];
                     TestBand1[10] = HoldGrid[LocalColumnPermutationTracker[1] + 9];
                     TestBand1[19] = HoldGrid[LocalColumnPermutationTracker[1] + 18];
                     TestBand1[28] = HoldGrid[LocalColumnPermutationTracker[1] + 27];
                     TestBand1[37] = HoldGrid[LocalColumnPermutationTracker[1] + 36];
                     TestBand1[46] = HoldGrid[LocalColumnPermutationTracker[1] + 45];
                     TestBand1[55] = HoldGrid[LocalColumnPermutationTracker[1] + 54];
                     TestBand1[64] = HoldGrid[LocalColumnPermutationTracker[1] + 63];
                     TestBand1[73] = HoldGrid[LocalColumnPermutationTracker[1] + 72];
                  }
                  if (LocalColumnPermutationTracker[2] != 2) // Column 3
                  {
                     TestBand1[2] = HoldGrid[LocalColumnPermutationTracker[2]];
                     TestBand1[11] = HoldGrid[LocalColumnPermutationTracker[2] + 9];
                     TestBand1[20] = HoldGrid[LocalColumnPermutationTracker[2] + 18];
                     TestBand1[29] = HoldGrid[LocalColumnPermutationTracker[2] + 27];
                     TestBand1[38] = HoldGrid[LocalColumnPermutationTracker[2] + 36];
                     TestBand1[47] = HoldGrid[LocalColumnPermutationTracker[2] + 45];
                     TestBand1[56] = HoldGrid[LocalColumnPermutationTracker[2] + 54];
                     TestBand1[65] = HoldGrid[LocalColumnPermutationTracker[2] + 63];
                     TestBand1[74] = HoldGrid[LocalColumnPermutationTracker[2] + 72];
                  }
                  if (LocalColumnPermutationTracker[3] != 3) // Column 4
                  {
                     TestBand1[3] = HoldGrid[LocalColumnPermutationTracker[3]];
                     TestBand1[12] = HoldGrid[LocalColumnPermutationTracker[3] + 9];
                     TestBand1[21] = HoldGrid[LocalColumnPermutationTracker[3] + 18];
                     TestBand1[30] = HoldGrid[LocalColumnPermutationTracker[3] + 27];
                     TestBand1[39] = HoldGrid[LocalColumnPermutationTracker[3] + 36];
                     TestBand1[48] = HoldGrid[LocalColumnPermutationTracker[3] + 45];
                     TestBand1[57] = HoldGrid[LocalColumnPermutationTracker[3] + 54];
                     TestBand1[66] = HoldGrid[LocalColumnPermutationTracker[3] + 63];
                     TestBand1[75] = HoldGrid[LocalColumnPermutationTracker[3] + 72];
                  }
                  if (LocalColumnPermutationTracker[4] != 4) // Column 5
                  {
                     TestBand1[4] = HoldGrid[LocalColumnPermutationTracker[4]];
                     TestBand1[13] = HoldGrid[LocalColumnPermutationTracker[4] + 9];
                     TestBand1[22] = HoldGrid[LocalColumnPermutationTracker[4] + 18];
                     TestBand1[31] = HoldGrid[LocalColumnPermutationTracker[4] + 27];
                     TestBand1[40] = HoldGrid[LocalColumnPermutationTracker[4] + 36];
                     TestBand1[49] = HoldGrid[LocalColumnPermutationTracker[4] + 45];
                     TestBand1[58] = HoldGrid[LocalColumnPermutationTracker[4] + 54];
                     TestBand1[67] = HoldGrid[LocalColumnPermutationTracker[4] + 63];
                     TestBand1[76] = HoldGrid[LocalColumnPermutationTracker[4] + 72];
                  }
                  if (LocalColumnPermutationTracker[5] != 5) // Column 6
                  {
                     TestBand1[5] = HoldGrid[LocalColumnPermutationTracker[5]];
                     TestBand1[14] = HoldGrid[LocalColumnPermutationTracker[5] + 9];
                     TestBand1[23] = HoldGrid[LocalColumnPermutationTracker[5] + 18];
                     TestBand1[32] = HoldGrid[LocalColumnPermutationTracker[5] + 27];
                     TestBand1[41] = HoldGrid[LocalColumnPermutationTracker[5] + 36];
                     TestBand1[50] = HoldGrid[LocalColumnPermutationTracker[5] + 45];
                     TestBand1[59] = HoldGrid[LocalColumnPermutationTracker[5] + 54];
                     TestBand1[68] = HoldGrid[LocalColumnPermutationTracker[5] + 63];
                     TestBand1[77] = HoldGrid[LocalColumnPermutationTracker[5] + 72];
                  }
                  if (LocalColumnPermutationTracker[6] != 6) // Column 7
                  {
                     TestBand1[6] = HoldGrid[LocalColumnPermutationTracker[6]];
                     TestBand1[15] = HoldGrid[LocalColumnPermutationTracker[6] + 9];
                     TestBand1[24] = HoldGrid[LocalColumnPermutationTracker[6] + 18];
                     TestBand1[33] = HoldGrid[LocalColumnPermutationTracker[6] + 27];
                     TestBand1[42] = HoldGrid[LocalColumnPermutationTracker[6] + 36];
                     TestBand1[51] = HoldGrid[LocalColumnPermutationTracker[6] + 45];
                     TestBand1[60] = HoldGrid[LocalColumnPermutationTracker[6] + 54];
                     TestBand1[69] = HoldGrid[LocalColumnPermutationTracker[6] + 63];
                     TestBand1[78] = HoldGrid[LocalColumnPermutationTracker[6] + 72];
                  }
                  if (LocalColumnPermutationTracker[7] != 7) // Column 8
                  {
                     TestBand1[7] = HoldGrid[LocalColumnPermutationTracker[7]];
                     TestBand1[16] = HoldGrid[LocalColumnPermutationTracker[7] + 9];
                     TestBand1[25] = HoldGrid[LocalColumnPermutationTracker[7] + 18];
                     TestBand1[34] = HoldGrid[LocalColumnPermutationTracker[7] + 27];
                     TestBand1[43] = HoldGrid[LocalColumnPermutationTracker[7] + 36];
                     TestBand1[52] = HoldGrid[LocalColumnPermutationTracker[7] + 45];
                     TestBand1[61] = HoldGrid[LocalColumnPermutationTracker[7] + 54];
                     TestBand1[70] = HoldGrid[LocalColumnPermutationTracker[7] + 63];
                     TestBand1[79] = HoldGrid[LocalColumnPermutationTracker[7] + 72];
                  }
                  if (LocalColumnPermutationTracker[8] != 8) // Column 9
                  {
                     TestBand1[8] = HoldGrid[LocalColumnPermutationTracker[8]];
                     TestBand1[17] = HoldGrid[LocalColumnPermutationTracker[8] + 9];
                     TestBand1[26] = HoldGrid[LocalColumnPermutationTracker[8] + 18];
                     TestBand1[35] = HoldGrid[LocalColumnPermutationTracker[8] + 27];
                     TestBand1[44] = HoldGrid[LocalColumnPermutationTracker[8] + 36];
                     TestBand1[53] = HoldGrid[LocalColumnPermutationTracker[8] + 45];
                     TestBand1[62] = HoldGrid[LocalColumnPermutationTracker[8] + 54];
                     TestBand1[71] = HoldGrid[LocalColumnPermutationTracker[8] + 63];
                     TestBand1[80] = HoldGrid[LocalColumnPermutationTracker[8] + 72];
                  }
               }

               for (i = 0; i <= 8; i++) // Mark Fixed Columns for the MinLexed Band1 (The Fixed Columns designation will be the same for all Candidates.)
               {
                  FixedColumns[i] = 0;
                  if (TestBand1[i] > 0 || TestBand1[i + 9] > 0 || TestBand1[i + 18] > 0)
                  {
                     FixedColumns[i] = 1;
                  }
               }
               if (FixedColumns[1] == 1 && FixedColumns[4] == 1 && FixedColumns[7] == 1)
               {
                  StillJustifyingSw = false;
                  StoppedJustifyingRow = 3;
               }
               else
               {
                  StillJustifyingSw = true;
               }

               MinLexCandidateSw = true;

               //  Identify candidates for Row 4.
               //  Row 4 Test: After right justification, test each row 4 to 9 with first non-zero digit position.
               //              And then, if more than one choice compare relabled digit values.
               Row4TestFirstNonZeroDigitPositionInRow = -1;
               FindFirstNonZeroDigitInRow(4, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[4], FirstNonZeroDigitRelabeled[4]);
               if (Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[4])
               {
                  Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[4];
                  StartEqualCheck = 4;
               }
               FindFirstNonZeroDigitInRow(5, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[5], FirstNonZeroDigitRelabeled[5]);
               if (Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[5])
               {
                  Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[5];
                  StartEqualCheck = 5;
               }
               FindFirstNonZeroDigitInRow(6, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[6], FirstNonZeroDigitRelabeled[6]);
               if (Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[6])
               {
                  Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[6];
                  StartEqualCheck = 6;
               }
               FindFirstNonZeroDigitInRow(7, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[7], FirstNonZeroDigitRelabeled[7]);
               if (Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[7])
               {
                  Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[7];
                  StartEqualCheck = 7;
               }
               FindFirstNonZeroDigitInRow(8, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[8], FirstNonZeroDigitRelabeled[8]);
               if (Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[8])
               {
                  Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[8];
                  StartEqualCheck = 8;
               }
               FindFirstNonZeroDigitInRow(9, StillJustifyingSw, TestBand1, FixedColumns, Row3DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[9], FirstNonZeroDigitRelabeled[9]);
               if (Row4TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[9])
               {
                  Row4TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[9];
                  StartEqualCheck = 9;
               }
               j = -1;
               CandidateFirstRelabeledDigit = 10;
               for (i = StartEqualCheck; i <= 9; i++)
               {
                  if (FirstNonZeroDigitPositionInRow[i] == Row4TestFirstNonZeroDigitPositionInRow)
                  {
                     j += 1;
                     Row4TestPositionalCandidateRow[j] = i;
                     iForHit[j] = i;
                     if (CandidateFirstRelabeledDigit > FirstNonZeroDigitRelabeled[i])
                     {
                        CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[i];
                     }
                  }
               }
               if (j > 0)
               {
                  Row4TestCandidateRowIx = -1;
                  for (i = 0; i <= j; i++)
                  {
                     if (CandidateFirstRelabeledDigit == FirstNonZeroDigitRelabeled[iForHit[i]])
                     {
                        Row4TestCandidateRowIx += 1;
                        Row4TestCandidateRow[Row4TestCandidateRowIx] = Row4TestPositionalCandidateRow[i];
                     }
                  }
               }
               else
               {
                  Row4TestCandidateRowIx = 0;
                  Row4TestCandidateRow[0] = Row4TestPositionalCandidateRow[0];
               }
               if (!CheckThisPassSw)
               {
                  if (Row4TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow[4])
                  {
                     Row4TestCandidateRowIx = -1;
                  }
                  else if (CandidateFirstRelabeledDigit < 10 && (Row4TestFirstNonZeroDigitPositionInRow == MinLexFirstNonZeroDigitPositionInRow[4] && CandidateFirstRelabeledDigit > MinLexGridLocal[27 + Row4TestFirstNonZeroDigitPositionInRow]))
                  {
                     Row4TestCandidateRowIx = -1;
                  }
               }
               if (Row4TestCandidateRowIx > 0)
               {
                  //Array.ConstrainedCopy(FixedColumns, 0, Row3FixedColumns, 0, 9) ' Save FixedColumns as of after Row 3.
                  for (z = 0; z <= 8; z++)
                  {
                     Row3FixedColumns[z] = FixedColumns[z];
                  }
               }
               for (int band2row4orderix = 0; band2row4orderix <= Row4TestCandidateRowIx; band2row4orderix++) // Process each row4 candidate.
               {
                  ArrayConstrainedCopy(TestBand1, 0, TestBand2, 0, 81);
                  switch (Row4TestCandidateRow[band2row4orderix]) // Move the next Row4Test candidate to row 4.
                  {
                     case 4: // Do nothing.
                        break;
                     case 5:
                        //Array.ConstrainedCopy(TestBand1, 36, TestBand2, 27, 9) ' Move row 5 to 4.
                        for (z = 27; z <= 35; z++)
                        {
                           TestBand2[z] = TestBand1[z + 9];
                        }
                        //Array.ConstrainedCopy(TestBand1, 27, TestBand2, 36, 9) ' Move row 4 to 5.
                        for (z = 27; z <= 35; z++)
                        {
                           TestBand2[z + 9] = TestBand1[z];
                        }
                        break;
                     case 6:
                        //Array.ConstrainedCopy(TestBand1, 45, TestBand2, 27, 9) ' Move row 6 to 4.
                        for (z = 27; z <= 35; z++)
                        {
                           TestBand2[z] = TestBand1[z + 18];
                        }
                        //Array.ConstrainedCopy(TestBand1, 27, TestBand2, 45, 9) ' Move row 4 to 6.
                        for (z = 27; z <= 35; z++)
                        {
                           TestBand2[z + 18] = TestBand1[z];
                        }
                        break;
                     case 7:
                        ArrayConstrainedCopy(TestBand1, 54, TestBand2, 27, 27); // Move band 3 to 2.
                        ArrayConstrainedCopy(TestBand1, 27, TestBand2, 54, 27); // Move band 2 to 3.
                        break;
                     case 8:
                        //Array.ConstrainedCopy(TestBand1, 63, TestBand2, 27, 18) ' Move rows 8 and 9 to 4 and 5.
                        for (z = 27; z <= 44; z++)
                        {
                           TestBand2[z] = TestBand1[z + 36];
                        }
                        //Array.ConstrainedCopy(TestBand1, 54, TestBand2, 45, 9) ' Move row 7 to 6.
                        for (z = 45; z <= 53; z++)
                        {
                           TestBand2[z] = TestBand1[z + 9];
                        }
                        ArrayConstrainedCopy(TestBand1, 27, TestBand2, 54, 27); // Move band 2 to band 3.
                        break;
                     case 9:
                        //Array.ConstrainedCopy(TestBand1, 72, TestBand2, 27, 9) ' Move row 9 to 4.
                        for (z = 27; z <= 35; z++)
                        {
                           TestBand2[z] = TestBand1[z + 45];
                        }
                        //Array.ConstrainedCopy(TestBand1, 54, TestBand2, 36, 18) ' Move rows 7 and 8 to 5 and 6.
                        for (z = 36; z <= 53; z++)
                        {
                           TestBand2[z] = TestBand1[z + 18];
                        }
                        ArrayConstrainedCopy(TestBand1, 27, TestBand2, 54, 27); // Move band 2 to band 3.
                        break;
                  }

                  FixedColumnsSavedAsOfRow4Sw = false;
                  if (StillJustifyingSw)
                  {
                     if (band2row4orderix > 0)
                     {
                        //Array.ConstrainedCopy(Row3FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 3 setting.
                        for (z = 0; z <= 8; z++)
                        {
                           FixedColumns[z] = Row3FixedColumns[z];
                        }
                     }
                     RightJustifyRow(4, TestBand2, StillJustifyingSw, FixedColumns, Row4StackPermutationCode, Row4ColumnPermutationCode);
                     if (!StillJustifyingSw)
                     {
                        StoppedJustifyingRow = 4;
                     }
                     if (Row4StackPermutationCode > 0 || Row4ColumnPermutationCode > 0)
                     {
                        //Array.ConstrainedCopy(FixedColumns, 0, Row4FixedColumns, 0, 9) ' Save FixedColumns as of after Row 4.
                        for (z = 0; z <= 8; z++)
                        {
                           Row4FixedColumns[z] = FixedColumns[z];
                        }
                        FixedColumnsSavedAsOfRow4Sw = true;
                     }
                  }
                  else
                  {
                     Row4StackPermutationCode = 0;
                     Row4ColumnPermutationCode = 0;
                  }
                  int tempVar5 = StackPermutations[ Row4StackPermutationCode];
                  for (int row4stackpermutationix = 0; row4stackpermutationix <= tempVar5; row4stackpermutationix++)
                  {
                     if (row4stackpermutationix > 0)
                     {
                        SwitchStacksXY(TestBand2, PermutationStackX[Row4StackPermutationCode][row4stackpermutationix], PermutationStackY[Row4StackPermutationCode][row4stackpermutationix]);
                     }
                     int tempVar6 = ColumnPermutations[ Row4ColumnPermutationCode];
                     for (int row4columnpermutationix = 0; row4columnpermutationix <= tempVar6; row4columnpermutationix++)
                     {
                        if (row4columnpermutationix > 0)
                        {
                           SwitchColumnsXY(TestBand2, PermutationColumnX[Row4ColumnPermutationCode][row4columnpermutationix], PermutationColumnY[Row4ColumnPermutationCode][row4columnpermutationix]);
                        }
                        //Array.ConstrainedCopy(Row3DigitAlreadyHitSw, 0, Row4DigitAlreadyHitSw, 0, 10)
                        for (z = 0; z <= 9; z++)
                        {
                           Row4DigitAlreadyHitSw[z] = Row3DigitAlreadyHitSw[z];
                        }
                        //Array.ConstrainedCopy(Row3DigitsRelabelWrk, 0, Row4DigitsRelabelWrk, 0, 10)
                        for (z = 0; z <= 9; z++)
                        {
                           Row4DigitsRelabelWrk[z] = Row3DigitsRelabelWrk[z];
                        }
                        Row4RelabelLastDigit = Row3RelabelLastDigit;
                        for (i = 27; i <= 35; i++) // Build Row4DigitsRelabelWrk and TestGridRelabeled for row4
                        {
                           if (TestBand2[i] > 0 && !Row4DigitAlreadyHitSw[TestBand2[i]])
                           {
                              Row4DigitAlreadyHitSw[TestBand2[i]] = true;
                              Row4RelabelLastDigit += 1;
                              Row4DigitsRelabelWrk[TestBand2[i]] = Row4RelabelLastDigit;
                           }
                           TestGridRelabeled[i] = Row4DigitsRelabelWrk[TestBand2[i]];
                        }
                        if (!CheckThisPassSw)
                        {
                           for (i = 27; i <= 35; i++)
                           {
                              if (TestGridRelabeled[i] > MinLexGridLocal[i]) // Check if Row4 is greater than MinLex.
                              {
                                 MinLexCandidateSw = false;
                                 break;
                              }
                              else if (TestGridRelabeled[i] < MinLexGridLocal[i])
                              {
                                 CheckThisPassSw = true;
                                 break;
                              }
                           }
                        }
                        if (StillJustifyingSw && (row4stackpermutationix > 0 || row4columnpermutationix > 0))
                        {
                           //Array.ConstrainedCopy(Row4FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 4 setting.
                           for (z = 0; z <= 8; z++)
                           {
                              FixedColumns[z] = Row4FixedColumns[z];
                           }
                        }
                        if (MinLexCandidateSw)
                        {
                           //  Identify candidates for Row 5.
                           //  Row 5 Test: After right justification, test each row 5 and 6 with first non-zero digit position and relabeled digit.
                           //              And then, if more than one choice compare relabeled digit values.
                           FindFirstNonZeroDigitInRow(5, StillJustifyingSw, TestBand2, FixedColumns, Row4DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[5], FirstNonZeroDigitRelabeled[5]);
                           FindFirstNonZeroDigitInRow(6, StillJustifyingSw, TestBand2, FixedColumns, Row4DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[6], FirstNonZeroDigitRelabeled[6]);

                           CandidateFirstRelabeledDigit = 99;
                           if (FirstNonZeroDigitPositionInRow[5] > FirstNonZeroDigitPositionInRow[6])
                           {
                              Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[5];
                              Row5TestCandidateRowIx = 0;
                              Row5TestCandidateRow[0] = 5;
                           }
                           else if (FirstNonZeroDigitPositionInRow[5] < FirstNonZeroDigitPositionInRow[6])
                           {
                              Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[6];
                              Row5TestCandidateRowIx = 0;
                              Row5TestCandidateRow[0] = 6;
                           }
                           else if (FirstNonZeroDigitRelabeled[5] < FirstNonZeroDigitRelabeled[6]) // Positions are equal.
                           {
                              Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[5];
                              Row5TestCandidateRowIx = 0;
                              Row5TestCandidateRow[0] = 5;
                              CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[5];
                           }
                           else if (FirstNonZeroDigitRelabeled[5] > FirstNonZeroDigitRelabeled[6])
                           {
                              Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[6];
                              Row5TestCandidateRowIx = 0;
                              Row5TestCandidateRow[0] = 6;
                              CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[6];
                           }
                           else
                           {
                              Row5TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[5];
                              Row5TestCandidateRowIx = 1; // This Case: Row 5 and 6 candidates are in the same position within the row and
                              Row5TestCandidateRow[0] = 5; // they both have the same relabeled value. For a valid Sudoku puzzle, that means the
                              Row5TestCandidateRow[1] = 6; // both have the default (unassigned) value of 10 or permutations produced the result, in which case they both need to be checked.
                              CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[5];
                           }
                           if (!CheckThisPassSw)
                           {
                              if (Row5TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow[5])
                              {
                                 Row5TestCandidateRowIx = -1;
                              }
                              else if (CandidateFirstRelabeledDigit < 10 && (Row5TestFirstNonZeroDigitPositionInRow == MinLexFirstNonZeroDigitPositionInRow[5] && CandidateFirstRelabeledDigit > MinLexGridLocal[36 + Row5TestFirstNonZeroDigitPositionInRow]))
                              {
                                 Row5TestCandidateRowIx = -1;
                              }
                           }

                           if (!FixedColumnsSavedAsOfRow4Sw && Row5TestCandidateRowIx > 0)
                           {
                              //Array.ConstrainedCopy(FixedColumns, 0, Row4FixedColumns, 0, 9)
                              for (z = 0; z <= 8; z++)
                              {
                                 Row4FixedColumns[z] = FixedColumns[z];
                              }
                              FixedColumnsSavedAsOfRow4Sw = true;
                           }
                           for (int band2rows5and6orderix = 0; band2rows5and6orderix <= Row5TestCandidateRowIx; band2rows5and6orderix++) // Process rows 5 and 6.
                           {
                              if (Row5TestCandidateRow[band2rows5and6orderix] == 6)
                              {
                                 //Array.ConstrainedCopy(TestBand2, 36, HoldRow, 0, 9) ' Switch rows 5 and 6.
                                 for (z = 0; z <= 8; z++)
                                 {
                                    HoldRow[z] = TestBand2[z + 36];
                                 }
                                 //Array.ConstrainedCopy(TestBand2, 45, TestBand2, 36, 9)
                                 for (z = 36; z <= 44; z++)
                                 {
                                    TestBand2[z] = TestBand2[z + 9];
                                 }
                                 //Array.ConstrainedCopy(HoldRow, 0, TestBand2, 45, 9)
                                 for (z = 0; z <= 8; z++)
                                 {
                                    TestBand2[z + 45] = HoldRow[z];
                                 }
                              }
                              if (StillJustifyingSw)
                              {
                                 if (band2rows5and6orderix > 0)
                                 {
                                    //Array.ConstrainedCopy(Row4FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 4 setting.
                                    for (z = 0; z <= 8; z++)
                                    {
                                       FixedColumns[z] = Row4FixedColumns[z];
                                    }
                                 }
                                 RightJustifyRow(5, TestBand2, StillJustifyingSw, FixedColumns, Row5StackPermutationCode, Row5ColumnPermutationCode);
                                 if (!StillJustifyingSw)
                                 {
                                    StoppedJustifyingRow = 5;
                                 }
                                 if (Row5StackPermutationCode > 0 || Row5ColumnPermutationCode > 0)
                                 {
                                    //Array.ConstrainedCopy(FixedColumns, 0, Row5FixedColumns, 0, 9) ' Save FixedColumns as of after Row 5.
                                    for (z = 0; z <= 8; z++)
                                    {
                                       Row5FixedColumns[z] = FixedColumns[z];
                                    }
                                 }
                              }
                              else
                              {
                                 Row5StackPermutationCode = 0;
                                 Row5ColumnPermutationCode = 0;
                              }

                              int tempVar7 = StackPermutations[ Row5StackPermutationCode];
                              for (int row5stackpermutationix = 0; row5stackpermutationix <= tempVar7; row5stackpermutationix++)
                              {
                                 if (row5stackpermutationix > 0)
                                 {
                                    SwitchStacksXY(TestBand2, PermutationStackX[Row5StackPermutationCode][row5stackpermutationix], PermutationStackY[Row5StackPermutationCode][row5stackpermutationix]);
                                 }
                                 int tempVar8 = ColumnPermutations[ Row5ColumnPermutationCode];
                                 for (int row5columnpermutationix = 0; row5columnpermutationix <= tempVar8; row5columnpermutationix++)
                                 {
                                    if (row5columnpermutationix > 0)
                                    {
                                       SwitchColumnsXY(TestBand2, PermutationColumnX[Row5ColumnPermutationCode][row5columnpermutationix], PermutationColumnY[Row5ColumnPermutationCode][row5columnpermutationix]);
                                    }
                                    //Array.ConstrainedCopy(Row4DigitAlreadyHitSw, 0, Row5DigitAlreadyHitSw, 0, 10)
                                    for (z = 0; z <= 9; z++)
                                    {
                                       Row5DigitAlreadyHitSw[z] = Row4DigitAlreadyHitSw[z];
                                    }
                                    //Array.ConstrainedCopy(Row4DigitsRelabelWrk, 0, Row5DigitsRelabelWrk, 0, 10)
                                    for (z = 0; z <= 9; z++)
                                    {
                                       Row5DigitsRelabelWrk[z] = Row4DigitsRelabelWrk[z];
                                    }
                                    Row5RelabelLastDigit = Row4RelabelLastDigit;
                                    for (i = 36; i <= 44; i++) // Build Row5DigitsRelabelWrk and TestGridRelabeled for row5
                                    {
                                       if (TestBand2[i] > 0 && !Row5DigitAlreadyHitSw[TestBand2[i]])
                                       {
                                          Row5DigitAlreadyHitSw[TestBand2[i]] = true;
                                          Row5RelabelLastDigit += 1;
                                          Row5DigitsRelabelWrk[TestBand2[i]] = Row5RelabelLastDigit;
                                       }
                                       TestGridRelabeled[i] = Row5DigitsRelabelWrk[TestBand2[i]];
                                    }
                                    if (!CheckThisPassSw)
                                    {
                                       for (i = 36; i <= 44; i++)
                                       {
                                          if (TestGridRelabeled[i] > MinLexGridLocal[i]) // Check if Row5 is greater than MinLex.
                                          {
                                             MinLexCandidateSw = false;
                                             break;
                                          }
                                          else if (TestGridRelabeled[i] < MinLexGridLocal[i])
                                          {
                                             CheckThisPassSw = true;
                                             break;
                                          }
                                       }
                                    }
                                    if (MinLexCandidateSw)
                                    {
                                       FixedColumnsSavedAsOfRow6Sw = false;
                                       if (StillJustifyingSw)
                                       {
                                          if (row5stackpermutationix > 0 || row5columnpermutationix > 0)
                                          {
                                             //Array.ConstrainedCopy(Row5FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 5 setting.
                                             for (z = 0; z <= 8; z++)
                                             {
                                                FixedColumns[z] = Row5FixedColumns[z];
                                             }
                                          }
                                          RightJustifyRow(6, TestBand2, StillJustifyingSw, FixedColumns, Row6StackPermutationCode, Row6ColumnPermutationCode);
                                          if (!StillJustifyingSw)
                                          {
                                             StoppedJustifyingRow = 6;
                                          }
                                          if (Row6StackPermutationCode > 0 || Row6ColumnPermutationCode > 0)
                                          {
                                             //Array.ConstrainedCopy(FixedColumns, 0, Row6FixedColumns, 0, 9) ' Save FixedColumns as of after Row 6.
                                             for (z = 0; z <= 8; z++)
                                             {
                                                Row6FixedColumns[z] = FixedColumns[z];
                                             }
                                             FixedColumnsSavedAsOfRow6Sw = true;
                                          }
                                       }
                                       else
                                       {
                                          Row6StackPermutationCode = 0;
                                          Row6ColumnPermutationCode = 0;
                                       }

                                       int tempVar9 = StackPermutations[ Row6StackPermutationCode];
                                       for (int row6stackpermutationix = 0; row6stackpermutationix <= tempVar9; row6stackpermutationix++)
                                       {
                                          if (row6stackpermutationix > 0)
                                          {
                                             SwitchStacksXY(TestBand2, PermutationStackX[Row6StackPermutationCode][row6stackpermutationix], PermutationStackY[Row6StackPermutationCode][row6stackpermutationix]);
                                          }
                                          int tempVar10 = ColumnPermutations[ Row6ColumnPermutationCode];
                                          for (int row6columnpermutationix = 0; row6columnpermutationix <= tempVar10; row6columnpermutationix++)
                                          {
                                             if (row6columnpermutationix > 0)
                                             {
                                                SwitchColumnsXY(TestBand2, PermutationColumnX[Row6ColumnPermutationCode][row6columnpermutationix], PermutationColumnY[Row6ColumnPermutationCode][row6columnpermutationix]);
                                             }
                                             //Array.ConstrainedCopy(Row5DigitAlreadyHitSw, 0, Row6DigitAlreadyHitSw, 0, 10)
                                             for (z = 0; z <= 9; z++)
                                             {
                                                Row6DigitAlreadyHitSw[z] = Row5DigitAlreadyHitSw[z];
                                             }
                                             //Array.ConstrainedCopy(Row5DigitsRelabelWrk, 0, Row6DigitsRelabelWrk, 0, 10)
                                             for (z = 0; z <= 9; z++)
                                             {
                                                Row6DigitsRelabelWrk[z] = Row5DigitsRelabelWrk[z];
                                             }
                                             Row6RelabelLastDigit = Row5RelabelLastDigit;
                                             for (i = 45; i <= 53; i++) // Build Row6DigitsRelabelWrk and TestGridRelabeled for row6
                                             {
                                                if (TestBand2[i] > 0 && !Row6DigitAlreadyHitSw[TestBand2[i]])
                                                {
                                                   Row6DigitAlreadyHitSw[TestBand2[i]] = true;
                                                   Row6RelabelLastDigit += 1;
                                                   Row6DigitsRelabelWrk[TestBand2[i]] = Row6RelabelLastDigit;
                                                }
                                                TestGridRelabeled[i] = Row6DigitsRelabelWrk[TestBand2[i]];
                                             }
                                             if (!CheckThisPassSw)
                                             {
                                                for (i = 45; i <= 53; i++)
                                                {
                                                   if (TestGridRelabeled[i] > MinLexGridLocal[i]) // Check if Row6 is greater than MinLex.
                                                   {
                                                      MinLexCandidateSw = false;
                                                      break;
                                                   }
                                                   else if (TestGridRelabeled[i] < MinLexGridLocal[i])
                                                   {
                                                      CheckThisPassSw = true;
                                                      break;
                                                   }
                                                }
                                             }
                                             if (MinLexCandidateSw) // Process Band3
                                             {
                                                // Identify candidates for Row 7.
                                                //        Row 7 Test: After right justification, test each row 7 to 9 first non-zero digit position and relabeled digit.
                                                //              And then, if more than one choice compare relabeled digit values.
                                                if (StillJustifyingSw && (row6stackpermutationix > 0 || row6columnpermutationix > 0))
                                                {
                                                   //Array.ConstrainedCopy(Row6FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 6 setting.
                                                   for (z = 0; z <= 8; z++)
                                                   {
                                                      FixedColumns[z] = Row6FixedColumns[z];
                                                   }
                                                }
                                                Row7TestFirstNonZeroDigitPositionInRow = -1;
                                                FindFirstNonZeroDigitInRow(7, StillJustifyingSw, TestBand2, FixedColumns, Row6DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[7], FirstNonZeroDigitRelabeled[7]);
                                                if (Row7TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[7])
                                                {
                                                   Row7TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[7];
                                                   StartEqualCheck = 7;
                                                }
                                                FindFirstNonZeroDigitInRow(8, StillJustifyingSw, TestBand2, FixedColumns, Row6DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[8], FirstNonZeroDigitRelabeled[8]);
                                                if (Row7TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[8])
                                                {
                                                   Row7TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[8];
                                                   StartEqualCheck = 8;
                                                }
                                                FindFirstNonZeroDigitInRow(9, StillJustifyingSw, TestBand2, FixedColumns, Row6DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[9], FirstNonZeroDigitRelabeled[9]);
                                                if (Row7TestFirstNonZeroDigitPositionInRow < FirstNonZeroDigitPositionInRow[9])
                                                {
                                                   Row7TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[9];
                                                   StartEqualCheck = 9;
                                                }
                                                j = -1;
                                                CandidateFirstRelabeledDigit = 10;
                                                for (i = StartEqualCheck; i <= 9; i++)
                                                {
                                                   if (FirstNonZeroDigitPositionInRow[i] == Row7TestFirstNonZeroDigitPositionInRow)
                                                   {
                                                      j += 1;
                                                      Row7TestPositionalCandidateRow[j] = i;
                                                      iForHit[j] = i;
                                                      if (CandidateFirstRelabeledDigit > FirstNonZeroDigitRelabeled[i])
                                                      {
                                                         CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[i];
                                                      }
                                                   }
                                                }
                                                if (j > 0)
                                                {
                                                   Row7TestCandidateRowIx = -1;
                                                   for (i = 0; i <= j; i++)
                                                   {
                                                      if (CandidateFirstRelabeledDigit == FirstNonZeroDigitRelabeled[iForHit[i]])
                                                      {
                                                         Row7TestCandidateRowIx += 1;
                                                         Row7TestCandidateRow[Row7TestCandidateRowIx] = Row7TestPositionalCandidateRow[i];
                                                      }
                                                   }
                                                }
                                                else
                                                {
                                                   Row7TestCandidateRowIx = 0;
                                                   Row7TestCandidateRow[0] = Row7TestPositionalCandidateRow[0];
                                                }
                                                if (!CheckThisPassSw)
                                                {
                                                   if (Row7TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow[7])
                                                   {
                                                      Row7TestCandidateRowIx = -1;
                                                   }
                                                   else if (CandidateFirstRelabeledDigit < 10 && (Row7TestFirstNonZeroDigitPositionInRow == MinLexFirstNonZeroDigitPositionInRow[7] && CandidateFirstRelabeledDigit > MinLexGridLocal[54 + Row7TestFirstNonZeroDigitPositionInRow]))
                                                   {
                                                      Row7TestCandidateRowIx = -1;
                                                   }
                                                }
                                                if (!FixedColumnsSavedAsOfRow6Sw && Row7TestCandidateRowIx > 0)
                                                {
                                                   //Array.ConstrainedCopy(FixedColumns, 0, Row6FixedColumns, 0, 9) ' Save FixedColumns as of after Row 6.
                                                   for (z = 0; z <= 8; z++)
                                                   {
                                                      Row6FixedColumns[z] = FixedColumns[z];
                                                   }
                                                   FixedColumnsSavedAsOfRow6Sw = true;
                                                }
                                                for (int band3row7orderix = 0; band3row7orderix <= Row7TestCandidateRowIx; band3row7orderix++) // Process each row 7 candidate.
                                                {
                                                   ArrayConstrainedCopy(TestBand2, 0, TestBand3, 0, 81);
                                                   switch (Row7TestCandidateRow[band3row7orderix]) // Move the next row 7 candidate to row 7.
                                                   {
                                                      case 8:
                                                         //Array.ConstrainedCopy(TestBand2, 54, TestBand3, 63, 9) ' Switch rows 7 and 8.
                                                         for (z = 54; z <= 62; z++)
                                                         {
                                                            TestBand3[z + 9] = TestBand2[z];
                                                         }
                                                         //Array.ConstrainedCopy(TestBand2, 63, TestBand3, 54, 9)
                                                         for (z = 54; z <= 62; z++)
                                                         {
                                                            TestBand3[z] = TestBand2[z + 9];
                                                         }
                                                         break;
                                                      case 9:
                                                         //Array.ConstrainedCopy(TestBand2, 54, TestBand3, 72, 9) ' Switch rows 7 and 9.
                                                         for (z = 54; z <= 62; z++)
                                                         {
                                                            TestBand3[z + 18] = TestBand2[z];
                                                         }
                                                         //Array.ConstrainedCopy(TestBand2, 72, TestBand3, 54, 9)
                                                         for (z = 54; z <= 62; z++)
                                                         {
                                                            TestBand3[z] = TestBand2[z + 18];
                                                         }
                                                         break;
                                                   }
                                                   FixedColumnsSavedAsOfRow7Sw = false;
                                                   if (StillJustifyingSw)
                                                   {
                                                      if (row6stackpermutationix > 0 || row6columnpermutationix > 0 || band3row7orderix > 0)
                                                      {
                                                         //Array.ConstrainedCopy(Row6FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 6 setting.
                                                         for (z = 0; z <= 8; z++)
                                                         {
                                                            FixedColumns[z] = Row6FixedColumns[z];
                                                         }
                                                      }
                                                      RightJustifyRow(7, TestBand3, StillJustifyingSw, FixedColumns, Row7StackPermutationCode, Row7ColumnPermutationCode);
                                                      if (!StillJustifyingSw)
                                                      {
                                                         StoppedJustifyingRow = 7;
                                                      }
                                                      if (Row7StackPermutationCode > 0 || Row7ColumnPermutationCode > 0)
                                                      {
                                                         //Array.ConstrainedCopy(FixedColumns, 0, Row7FixedColumns, 0, 9) ' Save FixedColumns as of after Row 7.
                                                         for (z = 0; z <= 8; z++)
                                                         {
                                                            Row7FixedColumns[z] = FixedColumns[z];
                                                         }
                                                         FixedColumnsSavedAsOfRow7Sw = true;
                                                      }
                                                   }
                                                   else
                                                   {
                                                      Row7StackPermutationCode = 0;
                                                      Row7ColumnPermutationCode = 0;
                                                   }

                                                   int tempVar11 = StackPermutations[ Row7StackPermutationCode];
                                                   for (int row7stackpermutationix = 0; row7stackpermutationix <= tempVar11; row7stackpermutationix++)
                                                   {
                                                      if (row7stackpermutationix > 0)
                                                      {
                                                         SwitchStacksXY(TestBand3, PermutationStackX[Row7StackPermutationCode][row7stackpermutationix], PermutationStackY[Row7StackPermutationCode][row7stackpermutationix]);
                                                      }
                                                      int tempVar12 = ColumnPermutations[ Row7ColumnPermutationCode];
                                                      for (int row7columnpermutationix = 0; row7columnpermutationix <= tempVar12; row7columnpermutationix++)
                                                      {
                                                         if (row7columnpermutationix > 0)
                                                         {
                                                            SwitchColumnsXY(TestBand3, PermutationColumnX[Row7ColumnPermutationCode][row7columnpermutationix], PermutationColumnY[Row7ColumnPermutationCode][row7columnpermutationix]);
                                                         }
                                                         //Array.ConstrainedCopy(Row6DigitAlreadyHitSw, 0, Row7DigitAlreadyHitSw, 0, 10)
                                                         for (z = 0; z <= 9; z++)
                                                         {
                                                            Row7DigitAlreadyHitSw[z] = Row6DigitAlreadyHitSw[z];
                                                         }
                                                         //Array.ConstrainedCopy(Row6DigitsRelabelWrk, 0, Row7DigitsRelabelWrk, 0, 10)
                                                         for (z = 0; z <= 9; z++)
                                                         {
                                                            Row7DigitsRelabelWrk[z] = Row6DigitsRelabelWrk[z];
                                                         }
                                                         Row7RelabelLastDigit = Row6RelabelLastDigit;
                                                         for (i = 54; i <= 62; i++) // Build Row7DigitsRelabelWrk and TestGridRelabeled for row7
                                                         {
                                                            if (TestBand3[i] > 0 && !Row7DigitAlreadyHitSw[TestBand3[i]])
                                                            {
                                                               Row7DigitAlreadyHitSw[TestBand3[i]] = true;
                                                               Row7RelabelLastDigit += 1;
                                                               Row7DigitsRelabelWrk[TestBand3[i]] = Row7RelabelLastDigit;
                                                            }
                                                            TestGridRelabeled[i] = Row7DigitsRelabelWrk[TestBand3[i]];
                                                         }
                                                         if (!CheckThisPassSw)
                                                         {
                                                            for (i = 54; i <= 62; i++)
                                                            {
                                                               if (TestGridRelabeled[i] > MinLexGridLocal[i]) // Check if Row7 is greater than MinLex.
                                                               {
                                                                  MinLexCandidateSw = false;
                                                                  break;
                                                               }
                                                               else if (TestGridRelabeled[i] < MinLexGridLocal[i])
                                                               {
                                                                  CheckThisPassSw = true;
                                                                  break;
                                                               }
                                                            }
                                                         }
                                                         if (MinLexCandidateSw)
                                                         {
                                                            //  Identify candidates for Row 8.
                                                            //  Row 8 Test: After right justification, test each row 8 and 9 with first non-zero digit position.
                                                            //              And then, if more than one choice compare relabeled digit values.
                                                            if (StillJustifyingSw && (row7stackpermutationix > 0 || row7columnpermutationix > 0))
                                                            {
                                                               //Array.ConstrainedCopy(Row7FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 7 setting.
                                                               for (z = 0; z <= 8; z++)
                                                               {
                                                                  FixedColumns[z] = Row7FixedColumns[z];
                                                               }
                                                            }
                                                            FindFirstNonZeroDigitInRow(8, StillJustifyingSw, TestBand3, FixedColumns, Row7DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[8], FirstNonZeroDigitRelabeled[8]);
                                                            FindFirstNonZeroDigitInRow(9, StillJustifyingSw, TestBand3, FixedColumns, Row7DigitsRelabelWrk, FirstNonZeroDigitPositionInRow[9], FirstNonZeroDigitRelabeled[9]);

                                                            CandidateFirstRelabeledDigit = 99;
                                                            if (FirstNonZeroDigitPositionInRow[8] > FirstNonZeroDigitPositionInRow[9])
                                                            {
                                                               Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[8];
                                                               Row8TestCandidateRowIx = 0;
                                                               Row8TestCandidateRow[0] = 8;
                                                            }
                                                            else if (FirstNonZeroDigitPositionInRow[8] < FirstNonZeroDigitPositionInRow[9])
                                                            {
                                                               Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[9];
                                                               Row8TestCandidateRowIx = 0;
                                                               Row8TestCandidateRow[0] = 9;
                                                            }
                                                            else if (FirstNonZeroDigitRelabeled[8] < FirstNonZeroDigitRelabeled[9]) // Positions are equal.
                                                            {
                                                               Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[8];
                                                               Row8TestCandidateRowIx = 0;
                                                               Row8TestCandidateRow[0] = 8;
                                                               CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[8];
                                                            }
                                                            else if (FirstNonZeroDigitRelabeled[8] > FirstNonZeroDigitRelabeled[9])
                                                            {
                                                               Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[9];
                                                               Row8TestCandidateRowIx = 0;
                                                               Row8TestCandidateRow[0] = 9;
                                                               CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[9];
                                                            }
                                                            else
                                                            {
                                                               Row8TestFirstNonZeroDigitPositionInRow = FirstNonZeroDigitPositionInRow[8];
                                                               Row8TestCandidateRowIx = 1; // This Case: Row 8 and 9 candidates are in the same position within the row and
                                                               Row8TestCandidateRow[0] = 8; // they both have the same relabeled value. For a valid Sudoku puzzle, that means the
                                                               Row8TestCandidateRow[1] = 9; // both have the default (unassigned) value of 10 or permutations produced the result, in which case they both need to be checked.
                                                               CandidateFirstRelabeledDigit = FirstNonZeroDigitRelabeled[8];
                                                            }
                                                            if (!CheckThisPassSw)
                                                            {
                                                               if (Row8TestFirstNonZeroDigitPositionInRow < MinLexFirstNonZeroDigitPositionInRow[8])
                                                               {
                                                                  Row8TestCandidateRowIx = -1;
                                                               }
                                                               else if (CandidateFirstRelabeledDigit < 10 && (Row8TestFirstNonZeroDigitPositionInRow == MinLexFirstNonZeroDigitPositionInRow[8] && CandidateFirstRelabeledDigit > MinLexGridLocal[63 + Row8TestFirstNonZeroDigitPositionInRow]))
                                                               {
                                                                  Row8TestCandidateRowIx = -1;
                                                               }
                                                            }

                                                            if (!FixedColumnsSavedAsOfRow7Sw && Row8TestCandidateRowIx > 0)
                                                            {
                                                               //Array.ConstrainedCopy(FixedColumns, 0, Row7FixedColumns, 0, 9)
                                                               for (z = 0; z <= 8; z++)
                                                               {
                                                                  Row7FixedColumns[z] = FixedColumns[z];
                                                               }
                                                               FixedColumnsSavedAsOfRow7Sw = true;
                                                            }
                                                            for (int band3rows8and9orderix = 0; band3rows8and9orderix <= Row8TestCandidateRowIx; band3rows8and9orderix++) // Process rows 8 and 9.
                                                            {
                                                               if (Row8TestCandidateRow[band3rows8and9orderix] == 9)
                                                               {
                                                                  //Array.ConstrainedCopy(TestBand3, 63, HoldRow, 0, 9) ' Switch rows 8 and 9.
                                                                  for (z = 0; z <= 8; z++)
                                                                  {
                                                                     HoldRow[z] = TestBand3[z + 63];
                                                                  }
                                                                  //Array.ConstrainedCopy(TestBand3, 72, TestBand3, 63, 9)
                                                                  for (z = 63; z <= 71; z++)
                                                                  {
                                                                     TestBand3[z] = TestBand3[z + 9];
                                                                  }
                                                                  //Array.ConstrainedCopy(HoldRow, 0, TestBand3, 72, 9)
                                                                  for (z = 0; z <= 8; z++)
                                                                  {
                                                                     TestBand3[z + 72] = HoldRow[z];
                                                                  }
                                                               }
                                                               if (StillJustifyingSw)
                                                               {
                                                                  if (band3rows8and9orderix > 0)
                                                                  {
                                                                     //Array.ConstrainedCopy(Row7FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row7 setting.
                                                                     for (z = 0; z <= 8; z++)
                                                                     {
                                                                        FixedColumns[z] = Row7FixedColumns[z];
                                                                     }
                                                                  }
                                                                  RightJustifyRow(8, TestBand3, StillJustifyingSw, FixedColumns, Row8StackPermutationCode, Row8ColumnPermutationCode);
                                                                  if (!StillJustifyingSw)
                                                                  {
                                                                     StoppedJustifyingRow = 8;
                                                                  }
                                                                  if (Row8StackPermutationCode > 0 || Row8ColumnPermutationCode > 0)
                                                                  {
                                                                     //Array.ConstrainedCopy(FixedColumns, 0, Row8FixedColumns, 0, 9) ' Save FixedColumns as of after Row 8.
                                                                     for (z = 0; z <= 8; z++)
                                                                     {
                                                                        Row8FixedColumns[z] = FixedColumns[z];
                                                                     }
                                                                  }
                                                               }
                                                               else
                                                               {
                                                                  Row8StackPermutationCode = 0;
                                                                  Row8ColumnPermutationCode = 0;
                                                               }

                                                               int tempVar13 = StackPermutations[ Row8StackPermutationCode];
                                                               for (int row8stackpermutationix = 0; row8stackpermutationix <= tempVar13; row8stackpermutationix++)
                                                               {
                                                                  if (row8stackpermutationix > 0)
                                                                  {
                                                                     SwitchStacksXY(TestBand3, PermutationStackX[Row8StackPermutationCode][row8stackpermutationix], PermutationStackY[Row8StackPermutationCode][row8stackpermutationix]);
                                                                  }
                                                                  int tempVar14 = ColumnPermutations[ Row8ColumnPermutationCode];
                                                                  for (int row8columnpermutationix = 0; row8columnpermutationix <= tempVar14; row8columnpermutationix++)
                                                                  {
                                                                     if (row8columnpermutationix > 0)
                                                                     {
                                                                        SwitchColumnsXY(TestBand3, PermutationColumnX[Row8ColumnPermutationCode][row8columnpermutationix], PermutationColumnY[Row8ColumnPermutationCode][row8columnpermutationix]);
                                                                     }
                                                                     //Array.ConstrainedCopy(Row7DigitAlreadyHitSw, 0, Row8DigitAlreadyHitSw, 0, 10)
                                                                     for (z = 0; z <= 9; z++)
                                                                     {
                                                                        Row8DigitAlreadyHitSw[z] = Row7DigitAlreadyHitSw[z];
                                                                     }
                                                                     //Array.ConstrainedCopy(Row7DigitsRelabelWrk, 0, Row8DigitsRelabelWrk, 0, 10)
                                                                     for (z = 0; z <= 9; z++)
                                                                     {
                                                                        Row8DigitsRelabelWrk[z] = Row7DigitsRelabelWrk[z];
                                                                     }
                                                                     Row8RelabelLastDigit = Row7RelabelLastDigit;
                                                                     for (i = 63; i <= 71; i++) // Build Row8DigitsRelabelWrk and TestGridRelabeled for row8
                                                                     {
                                                                        if (TestBand3[i] > 0 && !Row8DigitAlreadyHitSw[TestBand3[i]])
                                                                        {
                                                                           Row8DigitAlreadyHitSw[TestBand3[i]] = true;
                                                                           Row8RelabelLastDigit += 1;
                                                                           Row8DigitsRelabelWrk[TestBand3[i]] = Row8RelabelLastDigit;
                                                                        }
                                                                        TestGridRelabeled[i] = Row8DigitsRelabelWrk[TestBand3[i]];
                                                                     }
                                                                     if (!CheckThisPassSw)
                                                                     {
                                                                        for (i = 63; i <= 71; i++)
                                                                        {
                                                                           if (TestGridRelabeled[i] > MinLexGridLocal[i]) // Check if Row8 is greater than MinLex.
                                                                           {
                                                                              MinLexCandidateSw = false;
                                                                              break;
                                                                           }
                                                                           else if (TestGridRelabeled[i] < MinLexGridLocal[i])
                                                                           {
                                                                              CheckThisPassSw = true;
                                                                              break;
                                                                           }
                                                                        }
                                                                     }
                                                                     if (MinLexCandidateSw)
                                                                     {
                                                                        if (StillJustifyingSw)
                                                                        {
                                                                           if (row8stackpermutationix > 0 || row8columnpermutationix > 0)
                                                                           {
                                                                              //Array.ConstrainedCopy(Row8FixedColumns, 0, FixedColumns, 0, 9) ' Reset FixedColumns to after Row 8 setting.
                                                                              for (z = 0; z <= 8; z++)
                                                                              {
                                                                                 FixedColumns[z] = Row8FixedColumns[z];
                                                                              }
                                                                           }
                                                                           RightJustifyRow(9, TestBand3, StillJustifyingSw, FixedColumns, Row9StackPermutationCode, Row9ColumnPermutationCode);
                                                                           if (!StillJustifyingSw)
                                                                           {
                                                                              StoppedJustifyingRow = 9;
                                                                           }
                                                                        }
                                                                        else
                                                                        {
                                                                           Row9StackPermutationCode = 0;
                                                                           Row9ColumnPermutationCode = 0;
                                                                        }
                                                                        int tempVar15 = StackPermutations[ Row9StackPermutationCode];
                                                                        for (int row9stackpermutationix = 0; row9stackpermutationix <= tempVar15; row9stackpermutationix++)
                                                                        {
                                                                           if (row9stackpermutationix > 0)
                                                                           {
                                                                              SwitchStacksXY(TestBand3, PermutationStackX[Row9StackPermutationCode][row9stackpermutationix], PermutationStackY[Row9StackPermutationCode][row9stackpermutationix]);
                                                                           }
                                                                           int tempVar16 = ColumnPermutations[ Row9ColumnPermutationCode];
                                                                           for (int row9columnpermutationix = 0; row9columnpermutationix <= tempVar16; row9columnpermutationix++)
                                                                           {
                                                                              if (row9columnpermutationix > 0)
                                                                              {
                                                                                 SwitchColumnsXY(TestBand3, PermutationColumnX[Row9ColumnPermutationCode][row9columnpermutationix], PermutationColumnY[Row9ColumnPermutationCode][row9columnpermutationix]);
                                                                              }
                                                                              //Array.ConstrainedCopy(Row8DigitAlreadyHitSw, 0, Row9DigitAlreadyHitSw, 0, 10)
                                                                              for (z = 0; z <= 9; z++)
                                                                              {
                                                                                 Row9DigitAlreadyHitSw[z] = Row8DigitAlreadyHitSw[z];
                                                                              }
                                                                              //Array.ConstrainedCopy(Row8DigitsRelabelWrk, 0, Row9DigitsRelabelWrk, 0, 10)
                                                                              for (z = 0; z <= 9; z++)
                                                                              {
                                                                                 Row9DigitsRelabelWrk[z] = Row8DigitsRelabelWrk[z];
                                                                              }
                                                                              Row9RelabelLastDigit = Row8RelabelLastDigit;
                                                                              for (i = 72; i <= 80; i++) // Build Row9DigitsRelabelWrk and TestGridRelabeled for row9
                                                                              {
                                                                                 if (TestBand3[i] > 0 && !Row9DigitAlreadyHitSw[TestBand3[i]])
                                                                                 {
                                                                                    Row9DigitAlreadyHitSw[TestBand3[i]] = true;
                                                                                    Row9RelabelLastDigit += 1;
                                                                                    Row9DigitsRelabelWrk[TestBand3[i]] = Row9RelabelLastDigit;
                                                                                 }
                                                                                 TestGridRelabeled[i] = Row9DigitsRelabelWrk[TestBand3[i]];
                                                                              }
                                                                              if (!CheckThisPassSw)
                                                                              {
                                                                                 for (i = 72; i <= 80; i++)
                                                                                 {
                                                                                    if (TestGridRelabeled[i] > MinLexGridLocal[i]) // Check if Row9 is greater than MinLex.
                                                                                    {
                                                                                       MinLexCandidateSw = false;
                                                                                       break;
                                                                                    }
                                                                                    else if (TestGridRelabeled[i] < MinLexGridLocal[i])
                                                                                    {
                                                                                       CheckThisPassSw = true;
                                                                                       break;
                                                                                    }
                                                                                 }
                                                                                 if (i == 81) // If i = 81 then the TestBand1 matches the last found MinLex Puzzle candidate. If this happens for the actual MinLex Puzzle,
                                                                                 {
                                                                                    MinLexCandidateSw = false; // then the puzzle has multiple transformation paths that produce the minimal, thus the puzzle and its grid are automorphic.
                                                                                 }
                                                                              }
                                                                              if (MinLexCandidateSw)
                                                                              {
                                                                                 ArrayConstrainedCopy(TestGridRelabeled, 0, MinLexGridLocal, 0, 81);
                                                                                 CheckThisPassSw = false;
                                                                                 MinLexFirstNonZeroDigitPositionInRow[4] = Row4TestFirstNonZeroDigitPositionInRow;
                                                                                 MinLexFirstNonZeroDigitPositionInRow[5] = Row5TestFirstNonZeroDigitPositionInRow;
                                                                                 MinLexFirstNonZeroDigitPositionInRow[7] = Row7TestFirstNonZeroDigitPositionInRow;
                                                                                 MinLexFirstNonZeroDigitPositionInRow[8] = Row8TestFirstNonZeroDigitPositionInRow;
                                                                                 ArrayConstrainedCopy(TestGridRelabeled, 0, MinLexGridLocal, 0, 81);
                                                                              }
                                                                              MinLexCandidateSw = true;
                                                                           }
                                                                        }
                                                                        if (StoppedJustifyingRow == 9)
                                                                        {
                                                                           StoppedJustifyingRow = 0;
                                                                           StillJustifyingSw = true;
                                                                        }
                                                                     }
                                                                     MinLexCandidateSw = true;
                                                                  }
                                                               }
                                                               if (StoppedJustifyingRow == 8)
                                                               {
                                                                  StoppedJustifyingRow = 0;
                                                                  StillJustifyingSw = true;
                                                               }
                                                            }
                                                         }
                                                         MinLexCandidateSw = true;
                                                      }
                                                   }
                                                   if (StoppedJustifyingRow == 7)
                                                   {
                                                      StoppedJustifyingRow = 0;
                                                      StillJustifyingSw = true;
                                                   }
                                                }
                                             }
                                             MinLexCandidateSw = true;
                                          }
                                       }
                                       if (StoppedJustifyingRow == 6)
                                       {
                                          StoppedJustifyingRow = 0;
                                          StillJustifyingSw = true;
                                       }
                                    }
                                    MinLexCandidateSw = true;
                                 }
                              }
                              if (StoppedJustifyingRow == 5)
                              {
                                 StoppedJustifyingRow = 0;
                                 StillJustifyingSw = true;
                              }
                           }
                        }
                        MinLexCandidateSw = true;
                     }
                  }
                  if (StoppedJustifyingRow == 4)
                  {
                     StoppedJustifyingRow = 0;
                     StillJustifyingSw = true;
                  }
               }
            }
            FirstNonZeroRowCandidateIx += 1;
         } // Do While FirstNonZeroRowCandidateIx <= Step2Row1CandidateIx
      }
      else // Process full grid.
      {
         int HoldCell1 = 0;
         int HoldCell2 = 0;
         int HoldCell3 = 0;
         int TestGridRows[81];
         int TestGridCols[81];
         int TestGridBands[81];
         int TestGrid123[81];
         bool MinLexType456Sw = false;
         bool MinLexBandType456Sw[6] = { false, false, false, false, false, false};

         ArrayConstrainedCopy(MinLexFullGridLocalReset, 0, MinLexGridLocal, 0, 81);
         // Evaluate then InputGrid to determine if first 12 of Minlex is: 123456789456 or 123456789457
         // It will be "type456" if these exists a pairing of the same 3 digits in any two minirows of diferent rows of any band, otherwise it will be "type457".
         // For example if the following pattern is detected:   ------------------------------- where the digits 1, 2 and 4 occur in two different minirows,
         // then the minlex will be of type456.                 | 2  1  4 | x  x  x | x  x  x |
         //                                                     | x  x  x | 2  4  1 | x  x  x |
         //                                                     | x  x  x | x  x  x | x  x  x |
         //                                                     -------------------------------
         // It is sufficient to check just two cases, for example minirow1 of row1 with minirow2 of row2 and minirow3 of row2, to detect a type456 Minlex.
         for (i = 0; i <= 5; i++)
         {
            j = DigitsInMiniRowBit[i * 3][0] ^ DigitsInMiniRowBit[i * 3 + 1][1];
            k = DigitsInMiniRowBit[i * 3][0] ^ DigitsInMiniRowBit[i * 3 + 1][2];
            if (j == 0 || k == 0)
            {
               MinLexBandType456Sw[i] = true;
               if (!MinLexType456Sw)
               {
                  MinLexGridLocal[11] = 6;
                  MinLexGridLocal[12] = 7;
                  MinLexGridLocal[13] = 8;
                  MinLexGridLocal[14] = 9;
                  MinLexType456Sw = true;
               }
            }
         }
         if (MinLexType456Sw)
         {
            for (int LoopCount = 1; LoopCount <= 2; LoopCount++) // This loop executes two times, first finding the lexicographic minimum equivalent grid for the input grid - direct pass,
            {
               // ' then continuing with the transposed input grid - transposed pass.
               for (int band1order = 0; band1order <= 2; band1order++) // Move each band to the top in order
               {
                  if (MinLexBandType456Sw[band1order])
                  {
                     switch (band1order)
                     {
                        case 0:
                           ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81);
                           break;
                        case 1:
                           ArrayConstrainedCopy(InputGrid, 27, TestGridBands, 0, 27); // Switch Bands 1 & 2.
                           ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 27, 27);
                           ArrayConstrainedCopy(InputGrid, 54, TestGridBands, 54, 27);
                           break;
                        case 2:
                           ArrayConstrainedCopy(InputGrid, 54, TestGridBands, 0, 27); // Move Band3 to the top.
                           ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 27, 54); // Move Bands 1 & 2 to 2 & 3.
                           break;
                     }
                     for (int band1row1order = 0; band1row1order <= 2; band1row1order++) // Move each row in band1 to top
                     {
                        switch (band1row1order)
                        {
                           case 0:
                              ArrayConstrainedCopy(TestGridBands, 0, TestGridRows, 0, 81);
                              break;
                           case 1:
                              for (i = 0; i <= 8; i++) // Switch Rows 1 & 2.
                              {
                                 TestGridRows[i] = TestGridBands[9 + i];
                                 TestGridRows[9 + i] = TestGridBands[i];
                              }
                              ArrayConstrainedCopy(TestGridBands, 18, TestGridRows, 18, 63); // Copy Rows 3 to 9 to TestGridRows.
                              break;
                           case 2:
                              for (i = 0; i <= 8; i++)
                              {
                                 TestGridRows[i] = TestGridBands[18 + i]; // Move Rows 1, 2, 3 to Rows 2, 3, 1 respectively.
                                 TestGridRows[9 + i] = TestGridBands[i];
                                 TestGridRows[18 + i] = TestGridBands[9 + i];
                              }
                              ArrayConstrainedCopy(TestGridBands, 27, TestGridRows, 27, 54); // Copy Rows 4 to 9 to TestGridRows.
                              break;
                        }
                        for (int stackorder = 0; stackorder <= 5; stackorder++) // permute stacks
                        {
                           if (stackorder > 0)
                           {
                              SwitchStacksXY(TestGridRows, PermutationStackX[3][stackorder], PermutationStackY[3][stackorder]);
                           }
                           // Make sure row2 minirow1 containes the same digits as row1 minirow2.
                           if (!(TestGridRows[3] == TestGridRows[9] || TestGridRows[3] == TestGridRows[10] || TestGridRows[3] == TestGridRows[11]))
                           {
                              for (i = 9; i <= 17; i++) // Switch rows 2 & 3
                              {
                                 HoldCell1 = TestGridRows[i];
                                 TestGridRows[i] = TestGridRows[9 + i];
                                 TestGridRows[9 + i] = HoldCell1;
                              }
                           }
                           if (TestGridRows[3] == TestGridRows[9]) // Reorder row2 minirow1 to match row1 minirow2.
                           {
                              if (TestGridRows[4] == TestGridRows[11])
                              {
                                 SwitchColumns12(TestGridRows);
                              }
                           }
                           else if (TestGridRows[4] == TestGridRows[9])
                           {
                              if (TestGridRows[3] == TestGridRows[10])
                              {
                                 SwitchColumns01(TestGridRows);
                              }
                              else
                              {
                                 Switch3Columns201(TestGridRows); // Right circular shift: 123 ==> 312 (or 012 ==> 201 in 0-8 notation).
                              }
                           }
                           else if (TestGridRows[4] == TestGridRows[10])
                           {
                              SwitchColumns02(TestGridRows);
                           }
                           else
                           {
                              Switch3Columns120(TestGridRows); // Left circular shift: 123 ==> 231 (or 012 ==> 120 in 0-8 notation).
                           }
                           if (TestGridRows[12] == TestGridRows[6]) // Reorder row1 MiniRow3 to match row2 minirow2.
                           {
                              if (TestGridRows[13] == TestGridRows[8])
                              {
                                 SwitchColumns78(TestGridRows);
                              }
                           }
                           else if (TestGridRows[13] == TestGridRows[6])
                           {
                              if (TestGridRows[12] == TestGridRows[7])
                              {
                                 SwitchColumns67(TestGridRows);
                              }
                              else
                              {
                                 Switch3Columns867(TestGridRows); // Right circular shift: 789 ==> 978 (or 678 ==> 867 in 0-8 notation).
                              }
                           }
                           else if (TestGridRows[13] == TestGridRows[7])
                           {
                              SwitchColumns68(TestGridRows);
                           }
                           else
                           {
                              Switch3Columns786(TestGridRows); // Left circular shift: 789 ==> 897 (or 678 ==> 786 in 0-8 notation).
                           }
                           ArrayConstrainedCopy(TestGridRows, 0, TestGridCols, 0, 81);
                           for (int columnorder = 0; columnorder <= 5; columnorder++)
                           {
                              if (columnorder > 0) // Permute all three minirows in sync.
                              {
                                 SwitchColumnsXY(TestGridCols, PermutationStackX[3][columnorder], PermutationStackY[3][columnorder]);
                                 SwitchColumnsXY(TestGridCols, PermutationStackX[3][columnorder] + 3, PermutationStackY[3][columnorder] + 3);
                                 SwitchColumnsXY(TestGridCols, PermutationStackX[3][columnorder] + 6, PermutationStackY[3][columnorder] + 6);
                              }
                              for (i = 1; i <= 9; i++) // Build DigitsRelabelWrk
                              {
                                 DigitsRelabelWrk[TestGridCols[i - 1]] = i;
                              }
                              for (i = 0; i <= 26; i++) // ReLabel Band1 of TestGridCols into TestGrid123
                              {
                                 TestGrid123[i] = DigitsRelabelWrk[TestGridCols[i]];
                              }
                              MinLexCandidateSw = true;
                              for (i = 15; i <= 26; i++) // Note: row1 plus first six digits in row2 are 123456789456789, so start checking at row2 seventh digit.
                              {
                                 if (TestGrid123[i] > MinLexGridLocal[i])
                                 {
                                    MinLexCandidateSw = false;
                                    break;
                                 }
                                 else if (TestGrid123[i] < MinLexGridLocal[i])
                                 {
                                    break;
                                 }
                              }
                              if (MinLexCandidateSw)
                              {
                                 for (i = 27; i <= 80; i++) // ReLabel Band2 & Band3 of TestGridCols into TestGrid123
                                 {
                                    TestGrid123[i] = DigitsRelabelWrk[TestGridCols[i]];
                                 }
                                 if (TestGrid123[36] < TestGrid123[27] && TestGrid123[36] < TestGrid123[45]) // If row5 is smallest in band2, switch row4 and row5 and
                                 {
                                    for (i = 0; i <= 8; i++)
                                    {
                                       HoldCell1 = TestGrid123[27 + i];
                                       TestGrid123[27 + i] = TestGrid123[36 + i];
                                       TestGrid123[36 + i] = HoldCell1;
                                    }
                                 }
                                 else if (TestGrid123[45] < TestGrid123[27] && TestGrid123[45] < TestGrid123[36]) // ElseIf row6 is smallest in band2, switch row4 and row6 and ' (else, row6 is smallest in band2)
                                 {
                                    for (i = 0; i <= 8; i++)
                                    {
                                       HoldCell1 = TestGrid123[27 + i];
                                       TestGrid123[27 + i] = TestGrid123[45 + i];
                                       TestGrid123[45 + i] = HoldCell1;
                                    }
                                 }
                                 if (TestGrid123[45] < TestGrid123[36]) // If row6 is less than row5, then switch row5 and row6.
                                 {
                                    for (i = 0; i <= 8; i++)
                                    {
                                       HoldCell1 = TestGrid123[36 + i];
                                       TestGrid123[36 + i] = TestGrid123[45 + i];
                                       TestGrid123[45 + i] = HoldCell1;
                                    }
                                 }
                                 if (TestGrid123[63] < TestGrid123[54] && TestGrid123[63] < TestGrid123[72]) // If row8 is smallest in band3, then switch row7 and row8.
                                 {
                                    for (i = 0; i <= 8; i++)
                                    {
                                       HoldCell1 = TestGrid123[54 + i];
                                       TestGrid123[54 + i] = TestGrid123[63 + i];
                                       TestGrid123[63 + i] = HoldCell1;
                                    }
                                 }
                                 else if (TestGrid123[72] < TestGrid123[54] && TestGrid123[72] < TestGrid123[63]) // If row9 is smallest in band3, then switch row7 and row9.
                                 {
                                    for (i = 0; i <= 8; i++)
                                    {
                                       HoldCell1 = TestGrid123[54 + i];
                                       TestGrid123[54 + i] = TestGrid123[72 + i];
                                       TestGrid123[72 + i] = HoldCell1;
                                    }
                                 }
                                 if (TestGrid123[72] < TestGrid123[63]) // If row9 is less than row8, then switch row8 and row9.
                                 {
                                    for (i = 0; i <= 8; i++)
                                    {
                                       HoldCell1 = TestGrid123[63 + i];
                                       TestGrid123[63 + i] = TestGrid123[72 + i];
                                       TestGrid123[72 + i] = HoldCell1;
                                    }
                                 }
                                 if (TestGrid123[54] < TestGrid123[27]) // If row7 < row4 then switch band2 and band3.
                                 {
                                    for (i = 0; i <= 8; i++)
                                    {
                                       HoldCell1 = TestGrid123[27 + i];
                                       HoldCell2 = TestGrid123[36 + i];
                                       HoldCell3 = TestGrid123[45 + i];
                                       TestGrid123[27 + i] = TestGrid123[54 + i];
                                       TestGrid123[36 + i] = TestGrid123[63 + i];
                                       TestGrid123[45 + i] = TestGrid123[72 + i];
                                       TestGrid123[54 + i] = HoldCell1;
                                       TestGrid123[63 + i] = HoldCell2;
                                       TestGrid123[72 + i] = HoldCell3;
                                    }
                                 }
                                 // Check if TestGrid123 is smaller then current MinLexGridLocal - Note, row1 plus first two digits in row2 are always 12345678945.
                                 for (i = 11; i <= 80; i++)
                                 {
                                    if (TestGrid123[i] > MinLexGridLocal[i])
                                    {
                                       MinLexCandidateSw = false;
                                       break;
                                    }
                                    else if (TestGrid123[i] < MinLexGridLocal[i])
                                    {
                                       break;
                                    }
                                 }
                                 if (MinLexCandidateSw)
                                 {
                                    ArrayConstrainedCopy(TestGrid123, 0, MinLexGridLocal, 0, 81);
                                 }
                              }
                           }
                        }
                     }
                  } // If MinLexBandType456Sw Then
               }
               // transpose rows / columns, and continue to second loop.
               ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81); // TestGridBands is just used to do this transposition,
               for (i = 0; i <= 80; i++) // transpose rows to columns.
               {
                  InputGrid[i / 9 + (i % 9) * 9] = TestGridBands[i];
               }
               MinLexBandType456Sw[0] = MinLexBandType456Sw[3];
               MinLexBandType456Sw[1] = MinLexBandType456Sw[4];
               MinLexBandType456Sw[2] = MinLexBandType456Sw[5];
            }
         }
         else // Process full grid "457" case
         {
            for (int LoopCount = 1; LoopCount <= 2; LoopCount++) // This loop executes two times, first finding the lexicographic minimum equivalent grid for the input grid - direct pass,
            {
               // ' then continuing with the transposed input grid - transposed pass.
               for (int band1order = 0; band1order <= 2; band1order++) // Move each band to the top in order
               {
                  switch (band1order)
                  {
                     case 0:
                        ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81);
                        break;
                     case 1:
                        ArrayConstrainedCopy(InputGrid, 27, TestGridBands, 0, 27); // Switch Bands 1 & 2.
                        ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 27, 27);
                        ArrayConstrainedCopy(InputGrid, 54, TestGridBands, 54, 27);
                        break;
                     case 2:
                        ArrayConstrainedCopy(InputGrid, 54, TestGridBands, 0, 27); // Move Band3 to the top.
                        ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 27, 54); // Move Bands 1 & 2 to 2 & 3.
                        break;
                  }
                  for (int band1row1order = 0; band1row1order <= 2; band1row1order++) // Move each row in band1 to top
                  {
                     switch (band1row1order)
                     {
                        case 0:
                           ArrayConstrainedCopy(TestGridBands, 0, TestGridRows, 0, 81);
                           break;
                        case 1:
                           for (i = 0; i <= 8; i++) // Switch Rows 1 & 2.
                           {
                              TestGridRows[i] = TestGridBands[9 + i];
                              TestGridRows[9 + i] = TestGridBands[i];
                           }
                           ArrayConstrainedCopy(TestGridBands, 18, TestGridRows, 18, 63); // Copy Rows 3 to 9 to TestGridRows.
                           break;
                        case 2:
                           for (i = 0; i <= 8; i++)
                           {
                              TestGridRows[i] = TestGridBands[18 + i]; // Move Rows 1, 2, 3 to Rows 2, 3, 1 respectively.
                              TestGridRows[9 + i] = TestGridBands[i];
                              TestGridRows[18 + i] = TestGridBands[9 + i];
                           }
                           ArrayConstrainedCopy(TestGridBands, 27, TestGridRows, 27, 54); // Copy Rows 4 to 9 to TestGridRows.
                           break;
                     }
                     for (int stackorder = 0; stackorder <= 5; stackorder++) // permute stacks
                     {
                        if (stackorder > 0)
                        {
                           SwitchStacksXY(TestGridRows, PermutationStackX[3][stackorder], PermutationStackY[3][stackorder]);
                        }
                        // Make sure row2 minirow1 containes two of the same digits as row1 minirow2.
                        int MatchingDigit1 = 0;
                        int MatchingDigit1Position = 0;
                        int MatchingDigit2 = 0;
                        int MatchingDigit2Position = 0;
                        int MatchingDigitCount = 0;
                        if (TestGridRows[3] == TestGridRows[9] || TestGridRows[3] == TestGridRows[10] || TestGridRows[3] == TestGridRows[11])
                        {
                           MatchingDigit1 = TestGridRows[3];
                           MatchingDigit1Position = 3;
                           MatchingDigitCount = 1;
                           if (TestGridRows[4] == TestGridRows[9] || TestGridRows[4] == TestGridRows[10] || TestGridRows[4] == TestGridRows[11])
                           {
                              MatchingDigit2 = TestGridRows[4];
                              MatchingDigit2Position = 4;
                              MatchingDigitCount = 2;
                           }
                           else if (TestGridRows[5] == TestGridRows[9] || TestGridRows[5] == TestGridRows[10] || TestGridRows[5] == TestGridRows[11])
                           {
                              MatchingDigit2 = TestGridRows[5];
                              MatchingDigit2Position = 5;
                              MatchingDigitCount = 2;
                           }
                        }
                        else if (TestGridRows[4] == TestGridRows[9] || TestGridRows[4] == TestGridRows[10] || TestGridRows[4] == TestGridRows[11])
                        {
                           MatchingDigit1 = TestGridRows[4];
                           MatchingDigit1Position = 4;
                           MatchingDigitCount = 1;
                           if (TestGridRows[5] == TestGridRows[9] || TestGridRows[5] == TestGridRows[10] || TestGridRows[5] == TestGridRows[11])
                           {
                              MatchingDigit2 = TestGridRows[5];
                              MatchingDigit2Position = 5;
                              MatchingDigitCount = 2;
                           }
                        }
                        if (MatchingDigitCount < 2)
                        {
                           for (i = 9; i <= 17; i++) // Switch rows 2 & 3
                           {
                              HoldCell1 = TestGridRows[i];
                              TestGridRows[i] = TestGridRows[9 + i];
                              TestGridRows[9 + i] = HoldCell1;
                           }
                           if (TestGridRows[3] == TestGridRows[9] || TestGridRows[3] == TestGridRows[10] || TestGridRows[3] == TestGridRows[11])
                           {
                              MatchingDigit1 = TestGridRows[3];
                              MatchingDigit1Position = 3;
                              if (TestGridRows[4] == TestGridRows[9] || TestGridRows[4] == TestGridRows[10] || TestGridRows[4] == TestGridRows[11])
                              {
                                 MatchingDigit2 = TestGridRows[4];
                                 MatchingDigit2Position = 4;
                              }
                              else if (TestGridRows[5] == TestGridRows[9] || TestGridRows[5] == TestGridRows[10] || TestGridRows[5] == TestGridRows[11])
                              {
                                 MatchingDigit2 = TestGridRows[5];
                                 MatchingDigit2Position = 5;
                              }
                           }
                           else
                           {
                              MatchingDigit1 = TestGridRows[4];
                              MatchingDigit1Position = 4;
                              MatchingDigit2 = TestGridRows[5];
                              MatchingDigit2Position = 5;
                           }
                        }
                        if (MatchingDigit2Position == 5) // Adjust row1 minirow2 so the two matching digits are in positions 3 & 4 (0-8 notation)
                        {
                           if (MatchingDigit1Position == 4)
                           {
                              SwitchColumns35(TestGridRows);
                           }
                           else
                           {
                              SwitchColumns45(TestGridRows);
                           }
                        }
                        if (TestGridRows[3] == TestGridRows[9]) // Reorder two matching digits in row2 minirow1 to match row1 minirow2
                        {
                           if (TestGridRows[4] == TestGridRows[11])
                           {
                              SwitchColumns12(TestGridRows);
                           }
                        }
                        else if (TestGridRows[4] == TestGridRows[9])
                        {
                           if (TestGridRows[3] == TestGridRows[10])
                           {
                              SwitchColumns01(TestGridRows);
                           }
                           else
                           {
                              Switch3Columns201(TestGridRows); // Right circular shift: 123 ==> 312 (or 012 ==> 201 in 0-8 notation).
                           }
                        }
                        else if (TestGridRows[4] == TestGridRows[10])
                        {
                           SwitchColumns02(TestGridRows);
                        }
                        else
                        {
                           Switch3Columns120(TestGridRows); // Left circular shift: 123 ==> 231 (or 012 ==> 120 in 0-8 notation).
                        }
                        if (TestGridRows[11] == TestGridRows[7]) // Reorder row1 MiniRow3 position 6 to match the third cell in row2.
                        {
                           SwitchColumns67(TestGridRows);
                        }
                        else if (TestGridRows[11] == TestGridRows[8])
                        {
                           SwitchColumns68(TestGridRows);
                        }
                        // ' Reorder row1 MiniRow3 positions 7 and 8 to match the order of the matching digits in row2 minirow2.
                        if (TestGridRows[12] == TestGridRows[8] || (TestGridRows[13] == TestGridRows[8] && TestGridRows[12] != TestGridRows[7]))
                        {
                           SwitchColumns78(TestGridRows);
                        }
                        ArrayConstrainedCopy(TestGridRows, 0, TestGridCols, 0, 81);
                        for (int columnorder = 0; columnorder <= 1; columnorder++)
                        {
                           if (columnorder > 0) // Permute matching digits in first two minirows in sync.
                           {
                              SwitchColumns01(TestGridCols);
                              SwitchColumns34(TestGridCols);
                           }
                           // ' Reorder row1 MiniRow3 positions 7 and 8 to match the order of the matching digits in row2 minirow2.
                           if (TestGridRows[12] == TestGridRows[8] || (TestGridRows[13] == TestGridRows[8] && TestGridRows[12] != TestGridRows[7]))
                           {
                              SwitchColumns78(TestGridRows);
                           }
                           for (i = 1; i <= 9; i++) // Build DigitsRelabelWrk
                           {
                              DigitsRelabelWrk[TestGridCols[i - 1]] = i;
                           }
                           for (i = 0; i <= 26; i++) // ReLabel Band1 of TestGridCols into TestGrid123
                           {
                              TestGrid123[i] = DigitsRelabelWrk[TestGridCols[i]];
                           }
                           MinLexCandidateSw = true;
                           for (i = 12; i <= 26; i++) // Note: row1 plus first three digits in row2 are 123456789457, so start checking at row2 fourth digit.
                           {
                              if (TestGrid123[i] > MinLexGridLocal[i])
                              {
                                 MinLexCandidateSw = false;
                                 break;
                              }
                              else if (TestGrid123[i] < MinLexGridLocal[i])
                              {
                                 break;
                              }
                           }
                           if (MinLexCandidateSw)
                           {
                              for (i = 27; i <= 80; i++) // ReLabel Band2 & Band3 of TestGridCols into TestGrid123
                              {
                                 TestGrid123[i] = DigitsRelabelWrk[TestGridCols[i]];
                              }
                              if (TestGrid123[36] < TestGrid123[27] && TestGrid123[36] < TestGrid123[45]) // If row5 is smallest in band2, switch row4 and row5 and
                              {
                                 for (i = 0; i <= 8; i++)
                                 {
                                    HoldCell1 = TestGrid123[27 + i];
                                    TestGrid123[27 + i] = TestGrid123[36 + i];
                                    TestGrid123[36 + i] = HoldCell1;
                                 }
                              }
                              else if (TestGrid123[45] < TestGrid123[27] && TestGrid123[45] < TestGrid123[36]) // ElseIf row6 is smallest in band2, switch row4 and row6 and ' (else, row6 is smallest in band2)
                              {
                                 for (i = 0; i <= 8; i++)
                                 {
                                    HoldCell1 = TestGrid123[27 + i];
                                    TestGrid123[27 + i] = TestGrid123[45 + i];
                                    TestGrid123[45 + i] = HoldCell1;
                                 }
                              }
                              if (TestGrid123[45] < TestGrid123[36]) // If row6 is less than row5, then switch row5 and row6.
                              {
                                 for (i = 0; i <= 8; i++)
                                 {
                                    HoldCell1 = TestGrid123[36 + i];
                                    TestGrid123[36 + i] = TestGrid123[45 + i];
                                    TestGrid123[45 + i] = HoldCell1;
                                 }
                              }
                              if (TestGrid123[63] < TestGrid123[54] && TestGrid123[63] < TestGrid123[72]) // If row8 is smallest in band3, then switch row7 and row8.
                              {
                                 for (i = 0; i <= 8; i++)
                                 {
                                    HoldCell1 = TestGrid123[54 + i];
                                    TestGrid123[54 + i] = TestGrid123[63 + i];
                                    TestGrid123[63 + i] = HoldCell1;
                                 }
                              }
                              else if (TestGrid123[72] < TestGrid123[54] && TestGrid123[72] < TestGrid123[63]) // If row9 is smallest in band3, then switch row7 and row9.
                              {
                                 for (i = 0; i <= 8; i++)
                                 {
                                    HoldCell1 = TestGrid123[54 + i];
                                    TestGrid123[54 + i] = TestGrid123[72 + i];
                                    TestGrid123[72 + i] = HoldCell1;
                                 }
                              }
                              if (TestGrid123[72] < TestGrid123[63]) // If row9 is less than row8, then switch row8 and row9.
                              {
                                 for (i = 0; i <= 8; i++)
                                 {
                                    HoldCell1 = TestGrid123[63 + i];
                                    TestGrid123[63 + i] = TestGrid123[72 + i];
                                    TestGrid123[72 + i] = HoldCell1;
                                 }
                              }
                              if (TestGrid123[54] < TestGrid123[27]) // If row7 < row4 then switch band2 and band3.
                              {
                                 for (i = 0; i <= 8; i++)
                                 {
                                    HoldCell1 = TestGrid123[27 + i];
                                    HoldCell2 = TestGrid123[36 + i];
                                    HoldCell3 = TestGrid123[45 + i];
                                    TestGrid123[27 + i] = TestGrid123[54 + i];
                                    TestGrid123[36 + i] = TestGrid123[63 + i];
                                    TestGrid123[45 + i] = TestGrid123[72 + i];
                                    TestGrid123[54 + i] = HoldCell1;
                                    TestGrid123[63 + i] = HoldCell2;
                                    TestGrid123[72 + i] = HoldCell3;
                                 }
                              }
                              // Check if TestGrid123 is smaller then current MinLexGridLocal - Note, row1 plus first two digits in row2 are always 12345678945.
                              for (i = 11; i <= 80; i++)
                              {
                                 if (TestGrid123[i] > MinLexGridLocal[i])
                                 {
                                    MinLexCandidateSw = false;
                                    break;
                                 }
                                 else if (TestGrid123[i] < MinLexGridLocal[i])
                                 {
                                    break;
                                 }
                              }
                              if (MinLexCandidateSw)
                              {
                                 ArrayConstrainedCopy(TestGrid123, 0, MinLexGridLocal, 0, 81);
                              }
                           }
                        }
                     }
                  }
               }
               // transpose rows / columns, and continue to second loop.
               ArrayConstrainedCopy(InputGrid, 0, TestGridBands, 0, 81); // TestGridBands is just used to do this transposition,
               for (i = 0; i <= 80; i++) // transpose rows to columns.
               {
                  InputGrid[i / 9 + (i % 9) * 9] = TestGridBands[i];
               }
            }
         }
         // End ProcessFullGrid
      } // If CluesCount < 81
      zstart = ( ResultsBufferOffset + inputgridix) * 83;
      for (z = 0; z <= 80; z++) // On output, convert from integer array back to char ".' or 1 - 9.
      {
         MinLexBufferChr[zstart + z] = CharPeriodTO9[MinLexGridLocal[z]];
      }
      MinLexBufferChr[zstart + 81] = 0; // CarriageReturnChr;
      MinLexBufferChr[zstart + 82] = 0; // LineFeedChr;
   } // For inputgridix = 0 To InputLineCount - 1

   return 0; // successful run.
} // Public Sub MinLex9X9SR1

void RightJustifyRow(int Row, int Puzzle[], bool &StillJusfifyingSw, int FixedColumns[], int &StackPermutationCode, int &ColumnPermutationCode)
{
   // This routine is used for rows 4 to 9 to move non-zero digits to the right where possible and determine if stack or column permutations are required.
   // FixedColumns indicates which columns contain a non-zero digit above the provided row. It is reset based on the right justified digit positions of this row.

   int MiniRowCount[3] = { 0, 0, 0};
   int MiniRowPermutationCode[3] = { 0, 0, 0};

   int rowstart = (Row - 1) * 9;
   StackPermutationCode = 0;
   // Count non-zero digits in MiniRows.
   for (int i = 0; i <= 8; i++)
   {
      if (Puzzle[rowstart + i] > 0)
      {
         MiniRowCount[i / 3] += 1;
      }
   }

   // "right justify" MiniRows 1 andn 2.
   if (FixedColumns[5] == 0) // If minirow2 is not fixed, then
   {
      if (MiniRowCount[0] > MiniRowCount[1]) // If minirow1 is greater than minirow2, switch Stack1 and Stack2.
      {
         int temp = MiniRowCount[0]; MiniRowCount[0] = MiniRowCount[1]; MiniRowCount[1] = temp;
         SwitchStacks01(Puzzle);
      }
   }
   // "right justify" columns within MiniRows. Note: The FixedColumn indicators are right justified within MiniRows.
   // For each MiniRow (0 to 2 notation). Note: the comments below use "abc" notation for digits in the MiniRow.
   if (FixedColumns[1] == 0) // First MiniRow. If the "b" position is not fixed (then the "a" position would also not be fixed.), then
   {
      if (Puzzle[rowstart + 0] > 0 && Puzzle[rowstart + 1] == 0) // If "a" is non-zero and "b' is zero, then switch column "a" and column "b".
      {
         SwitchColumns01(Puzzle);
      }
      if (FixedColumns[2] == 0 && (Puzzle[rowstart + 1] > 0 && Puzzle[rowstart + 2] == 0)) // If "b" is non-zero and "c' is zero, then switch column "b" and column "c".
      {
         SwitchColumns12(Puzzle);
         if (Puzzle[rowstart + 0] > 0 && Puzzle[rowstart + 1] == 0) // Repeat first If statement above.
         {
            SwitchColumns01(Puzzle);
         }
      }
   }
   if (FixedColumns[4] == 0) // Second MiniRow.
   {
      if (Puzzle[rowstart + 3] > 0 && Puzzle[rowstart + 4] == 0)
      {
         SwitchColumns34(Puzzle);
      }
      if (FixedColumns[5] == 0 && (Puzzle[rowstart + 4] > 0 && Puzzle[rowstart + 5] == 0))
      {
         SwitchColumns45(Puzzle);
         if (Puzzle[rowstart + 3] > 0 && Puzzle[rowstart + 4] == 0)
         {
            SwitchColumns34(Puzzle);
         }
      }
   }
   if (FixedColumns[7] == 0) // Third MiniRow.
   {
      if (Puzzle[rowstart + 6] > 0 && Puzzle[rowstart + 7] == 0)
      {
         SwitchColumns67(Puzzle);
      }
   }

   // Set Stack PermutationCode
   if (FixedColumns[5] == 0 && (Puzzle[rowstart + 5]) > 0 && MiniRowCount[0] == MiniRowCount[1])
   {
      StackPermutationCode = 2;
   }

   // Set Column PermutationCode
   if (FixedColumns[2] == 0) // First MiniRow
   {
      if (Puzzle[rowstart] > 0)
      {
         MiniRowPermutationCode[0] = 3;
      }
      else if (Puzzle[rowstart + 1] > 0)
      {
         MiniRowPermutationCode[0] = 1;
      }
   }
   else if (FixedColumns[1] == 0 && (Puzzle[rowstart] > 0 && Puzzle[rowstart + 1] > 0))
   {
      MiniRowPermutationCode[0] = 2;
   }
   if (FixedColumns[5] == 0) // Second MiniRow
   {
      if (Puzzle[rowstart + 3] > 0)
      {
         MiniRowPermutationCode[1] = 3;
      }
      else if (Puzzle[rowstart + 4] > 0)
      {
         MiniRowPermutationCode[1] = 1;
      }
   }
   else if (FixedColumns[4] == 0 && (Puzzle[rowstart + 3] > 0 && Puzzle[rowstart + 4] > 0))
   {
      MiniRowPermutationCode[1] = 2;
   }
   if (FixedColumns[7] == 0 && (Puzzle[rowstart + 6] > 0 && Puzzle[rowstart + 7] > 0))
   {
      MiniRowPermutationCode[2] = 2;
   }
   ColumnPermutationCode = MiniRowPermutationCode[0] * 16 + MiniRowPermutationCode[1] * 4 + MiniRowPermutationCode[2];

   for (int i = 1; i <= 7; i++) // Mark Fixed Columns.
   {
      if (Puzzle[rowstart + i] > 0)
      {
         FixedColumns[i] = 1;
      }
   }

   if (FixedColumns[1] == 1 && FixedColumns[4] == 1 && FixedColumns[7] == 1)
   {
      StillJusfifyingSw = false;
   }
   else
   {
      StillJusfifyingSw = true;
   }
}

void FindFirstNonZeroDigitInRow(int Row, bool StillJustifyingSw, int Puzzle[], int FixedColumns[], int DigitsRelabelWrk[], int &FirstNonZeroDigitPositionInRow, int &FirstNonZeroDigitRelabeled)
{
   int LocalRow[9];
   int MiniRowCount[3] = { 0, 0, 0};
   int CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 0;
   int CandidateDigitsConsideringPossibleStackAndColumnPermutations[6] = { 0, 0, 0, 0, 0, 0};
   int FirstNonZeroDigit = 0;

// Array::ConstrainedCopy(Puzzle, (Row - 1) * 9, LocalRow, 0, 9);
   for (int i=0; i<9; i++ ) { LocalRow[ i] = Puzzle[ (Row-1)*9+i]; }

   if (StillJustifyingSw)
   {
      for (int i = 0; i <= 8; i++) // Count non-zero digits in minirow.
      {
         if (LocalRow[i] > 0)
         {
            MiniRowCount[i / 3] += 1;
         }
      }
      // "right justify" MiniRows in row. Note: The FixedColumns indicators are right justified within the row.
      if (FixedColumns[5] == 0) // If minirow2 is not fixed, then minirow1 is also not fixed.
      {
         if (MiniRowCount[0] > MiniRowCount[1]) // If minirow1 count is greater than minirow2, switch minirow1 and minirow2.
         {
            int temp = MiniRowCount[0]; MiniRowCount[0] = MiniRowCount[1]; MiniRowCount[1] = temp;
                temp = LocalRow[0]; LocalRow[0] = LocalRow[3]; LocalRow[3] = temp;
                temp = LocalRow[1]; LocalRow[1] = LocalRow[4]; LocalRow[4] = temp;
                temp = LocalRow[2]; LocalRow[2] = LocalRow[5]; LocalRow[5] = temp;
         }
      }

      // "right justify" columns in row within MiniRows. Note: The FixedColumn indicators are right justified within MiniRows.
      // For each MiniRow (0 to 2 notation). Note: the comments below use "abc" notation for digits in the MiniRow.
      if (FixedColumns[1] == 0) // First MiniRow ' If the "b" position is not fixed (then the "a" position would also not be fixed.), then
      {
         if (LocalRow[0] > 0 && LocalRow[1] == 0) // If "a" is non-zero and "b' is zero, then switch "a" and "b".
         {
            LocalRow[1] = LocalRow[0];
            LocalRow[0] = 0;
         }
         if (FixedColumns[2] == 0) // If "c" position is not fixed (then the "a" and "b" positions would also not be fixed.), then
         {
            if (LocalRow[1] > 0 && LocalRow[2] == 0) // If "b" is non-zero and "c' is zero, then switch "b" and "c".
            {
               LocalRow[2] = LocalRow[1];
               LocalRow[1] = 0;
               if (LocalRow[0] > 0 && LocalRow[1] == 0) // Repeat first If statement above.
               {
                  LocalRow[1] = LocalRow[0];
                  LocalRow[0] = 0;
               }
            }
         }
      }
      if (FixedColumns[4] == 0) // Second MiniRow
      {
         if (LocalRow[3] > 0 && LocalRow[4] == 0)
         {
            LocalRow[4] = LocalRow[3];
            LocalRow[3] = 0;
         }
         if (FixedColumns[5] == 0)
         {
            if (LocalRow[4] > 0 && LocalRow[5] == 0)
            {
               LocalRow[5] = LocalRow[4];
               LocalRow[4] = 0;
               if (LocalRow[3] > 0 && LocalRow[4] == 0)
               {
                  LocalRow[4] = LocalRow[3];
                  LocalRow[3] = 0;
               }
            }
         }
      }
      if (FixedColumns[7] == 0) // Third MiniRow
      {
         if (LocalRow[6] > 0 && LocalRow[7] == 0)
         {
            LocalRow[7] = LocalRow[6];
            LocalRow[6] = 0;
         }
      }

      FirstNonZeroDigitPositionInRow = 9; // Using 0-8 cell notation, 9 = empty row (all zeros).
      FirstNonZeroDigit = 10;
      for (int i = 0; i <= 8; i++) // Identify first non-zero digit position in row (if any).
      {
         if (LocalRow[i] > 0)
         {
            FirstNonZeroDigitPositionInRow = i;
            FirstNonZeroDigit = LocalRow[i];
            break;
         }
      }
//    CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 0;
      CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = FirstNonZeroDigit;
      if (FirstNonZeroDigitPositionInRow < 9) // If non-empty row, idenitfy candidate first digits in row considering possible permutations (assuming row position 9 is fixed).
      {
         // Idenitfy candidate first digits in row considering possible permutations (assuming row position 9 is fixed).
         if (FixedColumns[5] == 0 && MiniRowCount[0] > 0 && MiniRowCount[0] == MiniRowCount[1]) // Check for possible stack permutations.
         {
            switch (MiniRowCount[0])
            {
               case 1:
                  CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1;
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[2];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[5];
                  break;
               case 2:
                  CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 3;
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[1];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[2];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[2] = LocalRow[4];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[3] = LocalRow[5];
                  break;
               case 3:
                  CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 5;
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[0];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[1];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[2] = LocalRow[2];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[3] = LocalRow[3];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[4] = LocalRow[4];
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[5] = LocalRow[5];
                  break;
            }
         }
         else
         {
            switch (FirstNonZeroDigitPositionInRow) // Check for possible column permutations. Note: The FixedColumns indicators are right justified within MiniRows.
            {
               case 0:
                  if (LocalRow[1] > 0 && FixedColumns[1] == 0) // if FixedColumns(1) = 0 then FixedColumns(0) = 0 - ?? fix this.
                  {
                     if (LocalRow[2] > 0 && FixedColumns[2] == 0)
                     {
                        CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 2;
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[0];
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[1];
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[2] = LocalRow[2];
                     }
                     else
                     {
                        CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1;
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[0];
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[1];
                     }
                  }
                  break;
               case 1:
                  if (LocalRow[2] > 0 && FixedColumns[2] == 0)
                  {
                     CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1;
                     CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[1];
                     CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[2];
                  }
                  break;
               case 3:
                  if (LocalRow[4] > 0 && FixedColumns[4] == 0)
                  {
                     if (LocalRow[5] > 0 && FixedColumns[5] == 0)
                     {
                        CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 2;
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[3];
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[4];
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[2] = LocalRow[5];
                     }
                     else
                     {
                        CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1;
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[3];
                        CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[4];
                     }
                  }
                  break;
               case 4:
                  if (LocalRow[5] > 0 && FixedColumns[5] == 0)
                  {
                     CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1;
                     CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[4];
                     CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[5];
                  }
                  break;
               case 6:
                  if (LocalRow[7] > 0 && FixedColumns[7] == 0)
                  {
                     CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 1;
                     CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = LocalRow[6];
                     CandidateDigitsConsideringPossibleStackAndColumnPermutations[1] = LocalRow[7];
                  }
                  break;
               default:
                  CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx = 0;
                  CandidateDigitsConsideringPossibleStackAndColumnPermutations[0] = FirstNonZeroDigit;
                  break;
            }
         }

         FirstNonZeroDigitRelabeled = DigitsRelabelWrk[CandidateDigitsConsideringPossibleStackAndColumnPermutations[0]];
         for (int i = 1; i <= CandidateDigitsConsideringPossibleStackAndColumnPermutationsIx; i++)
         {
            if (FirstNonZeroDigitRelabeled > DigitsRelabelWrk[CandidateDigitsConsideringPossibleStackAndColumnPermutations[i]])
            {
               FirstNonZeroDigitRelabeled = DigitsRelabelWrk[CandidateDigitsConsideringPossibleStackAndColumnPermutations[i]];
            }
         }
      }
      else // empty row case
      {
         FirstNonZeroDigitRelabeled = 10;
      } //If FirstNonZeroDigitPositionInRow < 9 Then
   }
   else
   {
      FirstNonZeroDigitPositionInRow = 9; // Using 0-8 cell notation, 9 = empty row (all zeros).
      FirstNonZeroDigitRelabeled = 10;
      for (int i = 0; i <= 8; i++) // Identify first non-zero digit position in row (if any).
      {
         if (LocalRow[i] > 0)
         {
            FirstNonZeroDigitPositionInRow = i;
            FirstNonZeroDigitRelabeled = DigitsRelabelWrk[LocalRow[i]];
            break;
         }
      }
   }
}

void SwitchStacksXY( int Puzzle[], int stackXin, int stackYin)
{
   register int temp; register int stackX = stackXin*3; register int stackY = stackYin*3;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 7; stackY += 7;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp; stackX += 1; stackY += 1;
   temp = Puzzle[stackX]; Puzzle[stackX] = Puzzle[stackY]; Puzzle[stackY] = temp;
}

void SwitchStacks01( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 0]; Puzzle[ 0] = Puzzle[ 3]; Puzzle[ 3] = temp; // switch column 0-3
   temp = Puzzle[ 9]; Puzzle[ 9] = Puzzle[12]; Puzzle[12] = temp;
   temp = Puzzle[18]; Puzzle[18] = Puzzle[21]; Puzzle[21] = temp;
   temp = Puzzle[27]; Puzzle[27] = Puzzle[30]; Puzzle[30] = temp;
   temp = Puzzle[36]; Puzzle[36] = Puzzle[39]; Puzzle[39] = temp;
   temp = Puzzle[45]; Puzzle[45] = Puzzle[48]; Puzzle[48] = temp;
   temp = Puzzle[54]; Puzzle[54] = Puzzle[57]; Puzzle[57] = temp;
   temp = Puzzle[63]; Puzzle[63] = Puzzle[66]; Puzzle[66] = temp;
   temp = Puzzle[72]; Puzzle[72] = Puzzle[75]; Puzzle[75] = temp;
   temp = Puzzle[ 1]; Puzzle[ 1] = Puzzle[ 4]; Puzzle[ 4] = temp; // switch column 1-4
   temp = Puzzle[10]; Puzzle[10] = Puzzle[13]; Puzzle[13] = temp;
   temp = Puzzle[19]; Puzzle[19] = Puzzle[22]; Puzzle[22] = temp;
   temp = Puzzle[28]; Puzzle[28] = Puzzle[31]; Puzzle[31] = temp;
   temp = Puzzle[37]; Puzzle[37] = Puzzle[40]; Puzzle[40] = temp;
   temp = Puzzle[46]; Puzzle[46] = Puzzle[49]; Puzzle[49] = temp;
   temp = Puzzle[55]; Puzzle[55] = Puzzle[58]; Puzzle[58] = temp;
   temp = Puzzle[64]; Puzzle[64] = Puzzle[67]; Puzzle[67] = temp;
   temp = Puzzle[73]; Puzzle[73] = Puzzle[76]; Puzzle[76] = temp;
   temp = Puzzle[ 2]; Puzzle[ 2] = Puzzle[ 5]; Puzzle[ 5] = temp; // switch column 2-5
   temp = Puzzle[11]; Puzzle[11] = Puzzle[14]; Puzzle[14] = temp;
   temp = Puzzle[20]; Puzzle[20] = Puzzle[23]; Puzzle[23] = temp;
   temp = Puzzle[29]; Puzzle[29] = Puzzle[32]; Puzzle[32] = temp;
   temp = Puzzle[38]; Puzzle[38] = Puzzle[41]; Puzzle[41] = temp;
   temp = Puzzle[47]; Puzzle[47] = Puzzle[50]; Puzzle[50] = temp;
   temp = Puzzle[56]; Puzzle[56] = Puzzle[59]; Puzzle[59] = temp;
   temp = Puzzle[65]; Puzzle[65] = Puzzle[68]; Puzzle[68] = temp;
   temp = Puzzle[74]; Puzzle[74] = Puzzle[77]; Puzzle[77] = temp;
}

void SwitchStacks02( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 0]; Puzzle[ 0] = Puzzle[ 6]; Puzzle[ 6] = temp; // switch column 0-6
   temp = Puzzle[ 9]; Puzzle[ 9] = Puzzle[15]; Puzzle[15] = temp;
   temp = Puzzle[18]; Puzzle[18] = Puzzle[24]; Puzzle[24] = temp;
   temp = Puzzle[27]; Puzzle[27] = Puzzle[33]; Puzzle[33] = temp;
   temp = Puzzle[36]; Puzzle[36] = Puzzle[42]; Puzzle[42] = temp;
   temp = Puzzle[45]; Puzzle[45] = Puzzle[51]; Puzzle[51] = temp;
   temp = Puzzle[54]; Puzzle[54] = Puzzle[60]; Puzzle[60] = temp;
   temp = Puzzle[63]; Puzzle[63] = Puzzle[69]; Puzzle[69] = temp;
   temp = Puzzle[72]; Puzzle[72] = Puzzle[78]; Puzzle[78] = temp;
   temp = Puzzle[ 1]; Puzzle[ 1] = Puzzle[ 7]; Puzzle[ 7] = temp; // switch column 1-7
   temp = Puzzle[10]; Puzzle[10] = Puzzle[16]; Puzzle[16] = temp;
   temp = Puzzle[19]; Puzzle[19] = Puzzle[25]; Puzzle[25] = temp;
   temp = Puzzle[28]; Puzzle[28] = Puzzle[34]; Puzzle[34] = temp;
   temp = Puzzle[37]; Puzzle[37] = Puzzle[43]; Puzzle[43] = temp;
   temp = Puzzle[46]; Puzzle[46] = Puzzle[52]; Puzzle[52] = temp;
   temp = Puzzle[55]; Puzzle[55] = Puzzle[61]; Puzzle[61] = temp;
   temp = Puzzle[64]; Puzzle[64] = Puzzle[70]; Puzzle[70] = temp;
   temp = Puzzle[73]; Puzzle[73] = Puzzle[79]; Puzzle[79] = temp;
   temp = Puzzle[ 2]; Puzzle[ 2] = Puzzle[ 8]; Puzzle[ 8] = temp; // switch column 2-8
   temp = Puzzle[11]; Puzzle[11] = Puzzle[17]; Puzzle[17] = temp;
   temp = Puzzle[20]; Puzzle[20] = Puzzle[26]; Puzzle[26] = temp;
   temp = Puzzle[29]; Puzzle[29] = Puzzle[35]; Puzzle[35] = temp;
   temp = Puzzle[38]; Puzzle[38] = Puzzle[44]; Puzzle[44] = temp;
   temp = Puzzle[47]; Puzzle[47] = Puzzle[53]; Puzzle[53] = temp;
   temp = Puzzle[56]; Puzzle[56] = Puzzle[62]; Puzzle[62] = temp;
   temp = Puzzle[65]; Puzzle[65] = Puzzle[71]; Puzzle[71] = temp;
   temp = Puzzle[74]; Puzzle[74] = Puzzle[80]; Puzzle[80] = temp;
}

void SwitchStacks12( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 3]; Puzzle[ 3] = Puzzle[ 6]; Puzzle[ 6] = temp; // switch column 3-6
   temp = Puzzle[12]; Puzzle[12] = Puzzle[15]; Puzzle[15] = temp;
   temp = Puzzle[21]; Puzzle[21] = Puzzle[24]; Puzzle[24] = temp;
   temp = Puzzle[30]; Puzzle[30] = Puzzle[33]; Puzzle[33] = temp;
   temp = Puzzle[39]; Puzzle[39] = Puzzle[42]; Puzzle[42] = temp;
   temp = Puzzle[48]; Puzzle[48] = Puzzle[51]; Puzzle[51] = temp;
   temp = Puzzle[57]; Puzzle[57] = Puzzle[60]; Puzzle[60] = temp;
   temp = Puzzle[66]; Puzzle[66] = Puzzle[69]; Puzzle[69] = temp;
   temp = Puzzle[75]; Puzzle[75] = Puzzle[78]; Puzzle[78] = temp;
   temp = Puzzle[ 4]; Puzzle[ 4] = Puzzle[ 7]; Puzzle[ 7] = temp; // switch column 4-7
   temp = Puzzle[13]; Puzzle[13] = Puzzle[16]; Puzzle[16] = temp;
   temp = Puzzle[22]; Puzzle[22] = Puzzle[25]; Puzzle[25] = temp;
   temp = Puzzle[31]; Puzzle[31] = Puzzle[34]; Puzzle[34] = temp;
   temp = Puzzle[40]; Puzzle[40] = Puzzle[43]; Puzzle[43] = temp;
   temp = Puzzle[49]; Puzzle[49] = Puzzle[52]; Puzzle[52] = temp;
   temp = Puzzle[58]; Puzzle[58] = Puzzle[61]; Puzzle[61] = temp;
   temp = Puzzle[67]; Puzzle[67] = Puzzle[70]; Puzzle[70] = temp;
   temp = Puzzle[76]; Puzzle[76] = Puzzle[79]; Puzzle[79] = temp;
   temp = Puzzle[ 5]; Puzzle[ 5] = Puzzle[ 8]; Puzzle[ 8] = temp; // switch column 5-8
   temp = Puzzle[14]; Puzzle[14] = Puzzle[17]; Puzzle[17] = temp;
   temp = Puzzle[23]; Puzzle[23] = Puzzle[26]; Puzzle[26] = temp;
   temp = Puzzle[32]; Puzzle[32] = Puzzle[35]; Puzzle[35] = temp;
   temp = Puzzle[41]; Puzzle[41] = Puzzle[44]; Puzzle[44] = temp;
   temp = Puzzle[50]; Puzzle[50] = Puzzle[53]; Puzzle[53] = temp;
   temp = Puzzle[59]; Puzzle[59] = Puzzle[62]; Puzzle[62] = temp;
   temp = Puzzle[68]; Puzzle[68] = Puzzle[71]; Puzzle[71] = temp;
   temp = Puzzle[77]; Puzzle[77] = Puzzle[80]; Puzzle[80] = temp;
}

void Switch3Stacks120( int Puzzle[])   // 012 -> 120  rotate left
{
   register int temp;
   temp = Puzzle[ 0]; Puzzle[ 0] = Puzzle[ 3]; Puzzle[ 3] = Puzzle[ 6]; Puzzle[ 6] = temp;
   temp = Puzzle[ 1]; Puzzle[ 1] = Puzzle[ 4]; Puzzle[ 4] = Puzzle[ 7]; Puzzle[ 7] = temp;
   temp = Puzzle[ 2]; Puzzle[ 2] = Puzzle[ 5]; Puzzle[ 5] = Puzzle[ 8]; Puzzle[ 8] = temp;
   temp = Puzzle[ 9]; Puzzle[ 9] = Puzzle[12]; Puzzle[12] = Puzzle[15]; Puzzle[15] = temp;
   temp = Puzzle[10]; Puzzle[10] = Puzzle[13]; Puzzle[13] = Puzzle[16]; Puzzle[16] = temp;
   temp = Puzzle[11]; Puzzle[11] = Puzzle[14]; Puzzle[14] = Puzzle[17]; Puzzle[17] = temp;
   temp = Puzzle[18]; Puzzle[18] = Puzzle[21]; Puzzle[21] = Puzzle[24]; Puzzle[24] = temp;
   temp = Puzzle[19]; Puzzle[19] = Puzzle[22]; Puzzle[22] = Puzzle[25]; Puzzle[25] = temp;
   temp = Puzzle[20]; Puzzle[20] = Puzzle[23]; Puzzle[23] = Puzzle[26]; Puzzle[26] = temp;
   temp = Puzzle[27]; Puzzle[27] = Puzzle[30]; Puzzle[30] = Puzzle[33]; Puzzle[33] = temp;
   temp = Puzzle[28]; Puzzle[28] = Puzzle[31]; Puzzle[31] = Puzzle[34]; Puzzle[34] = temp;
   temp = Puzzle[29]; Puzzle[29] = Puzzle[32]; Puzzle[32] = Puzzle[35]; Puzzle[35] = temp;
   temp = Puzzle[36]; Puzzle[36] = Puzzle[39]; Puzzle[39] = Puzzle[42]; Puzzle[42] = temp;
   temp = Puzzle[37]; Puzzle[37] = Puzzle[40]; Puzzle[40] = Puzzle[43]; Puzzle[43] = temp;
   temp = Puzzle[38]; Puzzle[38] = Puzzle[41]; Puzzle[41] = Puzzle[44]; Puzzle[44] = temp;
   temp = Puzzle[45]; Puzzle[45] = Puzzle[48]; Puzzle[48] = Puzzle[51]; Puzzle[51] = temp;
   temp = Puzzle[46]; Puzzle[46] = Puzzle[49]; Puzzle[49] = Puzzle[52]; Puzzle[52] = temp;
   temp = Puzzle[47]; Puzzle[47] = Puzzle[50]; Puzzle[50] = Puzzle[53]; Puzzle[53] = temp;
   temp = Puzzle[54]; Puzzle[54] = Puzzle[57]; Puzzle[57] = Puzzle[60]; Puzzle[60] = temp;
   temp = Puzzle[55]; Puzzle[55] = Puzzle[58]; Puzzle[58] = Puzzle[61]; Puzzle[61] = temp;
   temp = Puzzle[56]; Puzzle[56] = Puzzle[59]; Puzzle[59] = Puzzle[62]; Puzzle[62] = temp;
   temp = Puzzle[63]; Puzzle[63] = Puzzle[66]; Puzzle[66] = Puzzle[69]; Puzzle[69] = temp;
   temp = Puzzle[64]; Puzzle[64] = Puzzle[67]; Puzzle[67] = Puzzle[70]; Puzzle[70] = temp;
   temp = Puzzle[65]; Puzzle[65] = Puzzle[68]; Puzzle[68] = Puzzle[71]; Puzzle[71] = temp;
   temp = Puzzle[72]; Puzzle[72] = Puzzle[75]; Puzzle[75] = Puzzle[78]; Puzzle[78] = temp;
   temp = Puzzle[73]; Puzzle[73] = Puzzle[76]; Puzzle[76] = Puzzle[79]; Puzzle[79] = temp;
   temp = Puzzle[74]; Puzzle[74] = Puzzle[77]; Puzzle[77] = Puzzle[80]; Puzzle[80] = temp;
}

void Switch3Stacks201( int Puzzle[])   // 012 -> 210  rotate right
{
   register int temp;
   temp = Puzzle[ 6]; Puzzle[ 6] = Puzzle[ 3]; Puzzle[ 3] = Puzzle[ 0]; Puzzle[ 0] = temp;
   temp = Puzzle[ 7]; Puzzle[ 7] = Puzzle[ 4]; Puzzle[ 4] = Puzzle[ 1]; Puzzle[ 1] = temp;
   temp = Puzzle[ 8]; Puzzle[ 8] = Puzzle[ 5]; Puzzle[ 5] = Puzzle[ 2]; Puzzle[ 2] = temp;
   temp = Puzzle[15]; Puzzle[15] = Puzzle[12]; Puzzle[12] = Puzzle[ 9]; Puzzle[ 9] = temp;
   temp = Puzzle[16]; Puzzle[16] = Puzzle[13]; Puzzle[13] = Puzzle[10]; Puzzle[10] = temp;
   temp = Puzzle[17]; Puzzle[17] = Puzzle[14]; Puzzle[14] = Puzzle[11]; Puzzle[11] = temp;
   temp = Puzzle[24]; Puzzle[24] = Puzzle[21]; Puzzle[21] = Puzzle[18]; Puzzle[18] = temp;
   temp = Puzzle[25]; Puzzle[25] = Puzzle[22]; Puzzle[22] = Puzzle[19]; Puzzle[19] = temp;
   temp = Puzzle[26]; Puzzle[26] = Puzzle[23]; Puzzle[23] = Puzzle[20]; Puzzle[20] = temp;
   temp = Puzzle[33]; Puzzle[33] = Puzzle[30]; Puzzle[30] = Puzzle[27]; Puzzle[27] = temp;
   temp = Puzzle[34]; Puzzle[34] = Puzzle[31]; Puzzle[31] = Puzzle[28]; Puzzle[28] = temp;
   temp = Puzzle[35]; Puzzle[35] = Puzzle[32]; Puzzle[32] = Puzzle[29]; Puzzle[29] = temp;
   temp = Puzzle[42]; Puzzle[42] = Puzzle[39]; Puzzle[39] = Puzzle[36]; Puzzle[36] = temp;
   temp = Puzzle[43]; Puzzle[43] = Puzzle[40]; Puzzle[40] = Puzzle[37]; Puzzle[37] = temp;
   temp = Puzzle[44]; Puzzle[44] = Puzzle[41]; Puzzle[41] = Puzzle[38]; Puzzle[38] = temp;
   temp = Puzzle[51]; Puzzle[51] = Puzzle[48]; Puzzle[48] = Puzzle[45]; Puzzle[45] = temp;
   temp = Puzzle[52]; Puzzle[52] = Puzzle[49]; Puzzle[49] = Puzzle[46]; Puzzle[46] = temp;
   temp = Puzzle[53]; Puzzle[53] = Puzzle[50]; Puzzle[50] = Puzzle[47]; Puzzle[47] = temp;
   temp = Puzzle[60]; Puzzle[60] = Puzzle[57]; Puzzle[57] = Puzzle[54]; Puzzle[54] = temp;
   temp = Puzzle[61]; Puzzle[61] = Puzzle[58]; Puzzle[58] = Puzzle[55]; Puzzle[55] = temp;
   temp = Puzzle[62]; Puzzle[62] = Puzzle[59]; Puzzle[59] = Puzzle[56]; Puzzle[56] = temp;
   temp = Puzzle[69]; Puzzle[69] = Puzzle[66]; Puzzle[66] = Puzzle[63]; Puzzle[63] = temp;
   temp = Puzzle[70]; Puzzle[70] = Puzzle[67]; Puzzle[67] = Puzzle[64]; Puzzle[64] = temp;
   temp = Puzzle[71]; Puzzle[71] = Puzzle[68]; Puzzle[68] = Puzzle[65]; Puzzle[65] = temp;
   temp = Puzzle[78]; Puzzle[78] = Puzzle[75]; Puzzle[75] = Puzzle[72]; Puzzle[72] = temp;
   temp = Puzzle[79]; Puzzle[79] = Puzzle[76]; Puzzle[76] = Puzzle[73]; Puzzle[73] = temp;
   temp = Puzzle[80]; Puzzle[80] = Puzzle[77]; Puzzle[77] = Puzzle[74]; Puzzle[74] = temp;
}

void SwitchColumnsXY( int Puzzle[], int columnX, int columnY)
{
   register int temp;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp; columnX += 9; columnY += 9;
   temp = Puzzle[columnX]; Puzzle[columnX] = Puzzle[columnY]; Puzzle[columnY] = temp;
}

void SwitchColumns01( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 0]; Puzzle[ 0] = Puzzle[ 1]; Puzzle[ 1] = temp;
   temp = Puzzle[ 9]; Puzzle[ 9] = Puzzle[10]; Puzzle[10] = temp;
   temp = Puzzle[18]; Puzzle[18] = Puzzle[19]; Puzzle[19] = temp;
   temp = Puzzle[27]; Puzzle[27] = Puzzle[28]; Puzzle[28] = temp;
   temp = Puzzle[36]; Puzzle[36] = Puzzle[37]; Puzzle[37] = temp;
   temp = Puzzle[45]; Puzzle[45] = Puzzle[46]; Puzzle[46] = temp;
   temp = Puzzle[54]; Puzzle[54] = Puzzle[55]; Puzzle[55] = temp;
   temp = Puzzle[63]; Puzzle[63] = Puzzle[64]; Puzzle[64] = temp;
   temp = Puzzle[72]; Puzzle[72] = Puzzle[73]; Puzzle[73] = temp;
}

void SwitchColumns02( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 0]; Puzzle[ 0] = Puzzle[ 2]; Puzzle[ 2] = temp;
   temp = Puzzle[ 9]; Puzzle[ 9] = Puzzle[11]; Puzzle[11] = temp;
   temp = Puzzle[18]; Puzzle[18] = Puzzle[20]; Puzzle[20] = temp;
   temp = Puzzle[27]; Puzzle[27] = Puzzle[29]; Puzzle[29] = temp;
   temp = Puzzle[36]; Puzzle[36] = Puzzle[38]; Puzzle[38] = temp;
   temp = Puzzle[45]; Puzzle[45] = Puzzle[47]; Puzzle[47] = temp;
   temp = Puzzle[54]; Puzzle[54] = Puzzle[56]; Puzzle[56] = temp;
   temp = Puzzle[63]; Puzzle[63] = Puzzle[65]; Puzzle[65] = temp;
   temp = Puzzle[72]; Puzzle[72] = Puzzle[74]; Puzzle[74] = temp;
}

void SwitchColumns12( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 1]; Puzzle[ 1] = Puzzle[ 2]; Puzzle[ 2] = temp;
   temp = Puzzle[10]; Puzzle[10] = Puzzle[11]; Puzzle[11] = temp;
   temp = Puzzle[19]; Puzzle[19] = Puzzle[20]; Puzzle[20] = temp;
   temp = Puzzle[28]; Puzzle[28] = Puzzle[29]; Puzzle[29] = temp;
   temp = Puzzle[37]; Puzzle[37] = Puzzle[38]; Puzzle[38] = temp;
   temp = Puzzle[46]; Puzzle[46] = Puzzle[47]; Puzzle[47] = temp;
   temp = Puzzle[55]; Puzzle[55] = Puzzle[56]; Puzzle[56] = temp;
   temp = Puzzle[64]; Puzzle[64] = Puzzle[65]; Puzzle[65] = temp;
   temp = Puzzle[73]; Puzzle[73] = Puzzle[74]; Puzzle[74] = temp;
}

void SwitchColumns34( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 3]; Puzzle[ 3] = Puzzle[ 4]; Puzzle[ 4] = temp;
   temp = Puzzle[12]; Puzzle[12] = Puzzle[13]; Puzzle[13] = temp;
   temp = Puzzle[21]; Puzzle[21] = Puzzle[22]; Puzzle[22] = temp;
   temp = Puzzle[30]; Puzzle[30] = Puzzle[31]; Puzzle[31] = temp;
   temp = Puzzle[39]; Puzzle[39] = Puzzle[40]; Puzzle[40] = temp;
   temp = Puzzle[48]; Puzzle[48] = Puzzle[49]; Puzzle[49] = temp;
   temp = Puzzle[57]; Puzzle[57] = Puzzle[58]; Puzzle[58] = temp;
   temp = Puzzle[66]; Puzzle[66] = Puzzle[67]; Puzzle[67] = temp;
   temp = Puzzle[75]; Puzzle[75] = Puzzle[76]; Puzzle[76] = temp;
}

void SwitchColumns35( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 3]; Puzzle[ 3] = Puzzle[ 5]; Puzzle[ 5] = temp;
   temp = Puzzle[12]; Puzzle[12] = Puzzle[14]; Puzzle[14] = temp;
   temp = Puzzle[21]; Puzzle[21] = Puzzle[23]; Puzzle[23] = temp;
   temp = Puzzle[30]; Puzzle[30] = Puzzle[32]; Puzzle[32] = temp;
   temp = Puzzle[39]; Puzzle[39] = Puzzle[41]; Puzzle[41] = temp;
   temp = Puzzle[48]; Puzzle[48] = Puzzle[50]; Puzzle[50] = temp;
   temp = Puzzle[57]; Puzzle[57] = Puzzle[59]; Puzzle[59] = temp;
   temp = Puzzle[66]; Puzzle[66] = Puzzle[68]; Puzzle[68] = temp;
   temp = Puzzle[75]; Puzzle[75] = Puzzle[77]; Puzzle[77] = temp;
}

void SwitchColumns45( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 4]; Puzzle[ 4] = Puzzle[ 5]; Puzzle[ 5] = temp;
   temp = Puzzle[13]; Puzzle[13] = Puzzle[14]; Puzzle[14] = temp;
   temp = Puzzle[22]; Puzzle[22] = Puzzle[23]; Puzzle[23] = temp;
   temp = Puzzle[31]; Puzzle[31] = Puzzle[32]; Puzzle[32] = temp;
   temp = Puzzle[40]; Puzzle[40] = Puzzle[41]; Puzzle[41] = temp;
   temp = Puzzle[49]; Puzzle[49] = Puzzle[50]; Puzzle[50] = temp;
   temp = Puzzle[58]; Puzzle[58] = Puzzle[59]; Puzzle[59] = temp;
   temp = Puzzle[67]; Puzzle[67] = Puzzle[68]; Puzzle[68] = temp;
   temp = Puzzle[76]; Puzzle[76] = Puzzle[77]; Puzzle[77] = temp;
}

void SwitchColumns67( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 6]; Puzzle[ 6] = Puzzle[ 7]; Puzzle[ 7] = temp;
   temp = Puzzle[15]; Puzzle[15] = Puzzle[16]; Puzzle[16] = temp;
   temp = Puzzle[24]; Puzzle[24] = Puzzle[25]; Puzzle[25] = temp;
   temp = Puzzle[33]; Puzzle[33] = Puzzle[34]; Puzzle[34] = temp;
   temp = Puzzle[42]; Puzzle[42] = Puzzle[43]; Puzzle[43] = temp;
   temp = Puzzle[51]; Puzzle[51] = Puzzle[52]; Puzzle[52] = temp;
   temp = Puzzle[60]; Puzzle[60] = Puzzle[61]; Puzzle[61] = temp;
   temp = Puzzle[69]; Puzzle[69] = Puzzle[70]; Puzzle[70] = temp;
   temp = Puzzle[78]; Puzzle[78] = Puzzle[79]; Puzzle[79] = temp;
}

void SwitchColumns68( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 6]; Puzzle[ 6] = Puzzle[ 8]; Puzzle[ 8] = temp;
   temp = Puzzle[15]; Puzzle[15] = Puzzle[17]; Puzzle[17] = temp;
   temp = Puzzle[24]; Puzzle[24] = Puzzle[26]; Puzzle[26] = temp;
   temp = Puzzle[33]; Puzzle[33] = Puzzle[35]; Puzzle[35] = temp;
   temp = Puzzle[42]; Puzzle[42] = Puzzle[44]; Puzzle[44] = temp;
   temp = Puzzle[51]; Puzzle[51] = Puzzle[53]; Puzzle[53] = temp;
   temp = Puzzle[60]; Puzzle[60] = Puzzle[62]; Puzzle[62] = temp;
   temp = Puzzle[69]; Puzzle[69] = Puzzle[71]; Puzzle[71] = temp;
   temp = Puzzle[78]; Puzzle[78] = Puzzle[80]; Puzzle[80] = temp;
}

void SwitchColumns78( int Puzzle[])
{
   register int temp;
   temp = Puzzle[ 7]; Puzzle[ 7] = Puzzle[ 8]; Puzzle[ 8] = temp;
   temp = Puzzle[16]; Puzzle[16] = Puzzle[17]; Puzzle[17] = temp;
   temp = Puzzle[25]; Puzzle[25] = Puzzle[26]; Puzzle[26] = temp;
   temp = Puzzle[34]; Puzzle[34] = Puzzle[35]; Puzzle[35] = temp;
   temp = Puzzle[43]; Puzzle[43] = Puzzle[44]; Puzzle[44] = temp;
   temp = Puzzle[52]; Puzzle[52] = Puzzle[53]; Puzzle[53] = temp;
   temp = Puzzle[61]; Puzzle[61] = Puzzle[62]; Puzzle[62] = temp;
   temp = Puzzle[70]; Puzzle[70] = Puzzle[71]; Puzzle[71] = temp;
   temp = Puzzle[79]; Puzzle[79] = Puzzle[80]; Puzzle[80] = temp;
}

void Switch3Columns120( int Puzzle[])  // 012 -> 120  rotate left
{
   register int temp;
   temp = Puzzle[ 0]; Puzzle[ 0] = Puzzle[ 1]; Puzzle[ 1] = Puzzle[ 2]; Puzzle[ 2] = temp;
   temp = Puzzle[ 9]; Puzzle[ 9] = Puzzle[10]; Puzzle[10] = Puzzle[11]; Puzzle[11] = temp;
   temp = Puzzle[18]; Puzzle[18] = Puzzle[19]; Puzzle[19] = Puzzle[20]; Puzzle[20] = temp;
   temp = Puzzle[27]; Puzzle[27] = Puzzle[28]; Puzzle[28] = Puzzle[29]; Puzzle[29] = temp;
   temp = Puzzle[36]; Puzzle[36] = Puzzle[37]; Puzzle[37] = Puzzle[38]; Puzzle[38] = temp;
   temp = Puzzle[45]; Puzzle[45] = Puzzle[46]; Puzzle[46] = Puzzle[47]; Puzzle[47] = temp;
   temp = Puzzle[54]; Puzzle[54] = Puzzle[55]; Puzzle[55] = Puzzle[56]; Puzzle[56] = temp;
   temp = Puzzle[63]; Puzzle[63] = Puzzle[64]; Puzzle[64] = Puzzle[65]; Puzzle[65] = temp;
   temp = Puzzle[72]; Puzzle[72] = Puzzle[73]; Puzzle[73] = Puzzle[74]; Puzzle[74] = temp;
}

void Switch3Columns201( int Puzzle[])  // 012 -> 201  rotate right
{
   register int temp;
   temp = Puzzle[ 2]; Puzzle[ 2] = Puzzle[ 1]; Puzzle[ 1] = Puzzle[ 0]; Puzzle[ 0] = temp;
   temp = Puzzle[11]; Puzzle[11] = Puzzle[10]; Puzzle[10] = Puzzle[ 9]; Puzzle[ 9] = temp;
   temp = Puzzle[20]; Puzzle[20] = Puzzle[19]; Puzzle[19] = Puzzle[18]; Puzzle[18] = temp;
   temp = Puzzle[29]; Puzzle[29] = Puzzle[28]; Puzzle[28] = Puzzle[27]; Puzzle[27] = temp;
   temp = Puzzle[38]; Puzzle[38] = Puzzle[37]; Puzzle[37] = Puzzle[36]; Puzzle[36] = temp;
   temp = Puzzle[47]; Puzzle[47] = Puzzle[46]; Puzzle[46] = Puzzle[45]; Puzzle[45] = temp;
   temp = Puzzle[56]; Puzzle[56] = Puzzle[55]; Puzzle[55] = Puzzle[54]; Puzzle[54] = temp;
   temp = Puzzle[65]; Puzzle[65] = Puzzle[64]; Puzzle[64] = Puzzle[63]; Puzzle[63] = temp;
   temp = Puzzle[74]; Puzzle[74] = Puzzle[73]; Puzzle[73] = Puzzle[72]; Puzzle[72] = temp;
}

void Switch3Columns786( int Puzzle[])  // 678 -> 786  rotate left
{
   register int temp;
   temp = Puzzle[ 6]; Puzzle[ 6] = Puzzle[ 7]; Puzzle[ 7] = Puzzle[ 8]; Puzzle[ 8] = temp;
   temp = Puzzle[15]; Puzzle[15] = Puzzle[16]; Puzzle[16] = Puzzle[17]; Puzzle[17] = temp;
   temp = Puzzle[24]; Puzzle[24] = Puzzle[25]; Puzzle[25] = Puzzle[26]; Puzzle[26] = temp;
   temp = Puzzle[33]; Puzzle[33] = Puzzle[34]; Puzzle[34] = Puzzle[35]; Puzzle[35] = temp;
   temp = Puzzle[42]; Puzzle[42] = Puzzle[43]; Puzzle[43] = Puzzle[44]; Puzzle[44] = temp;
   temp = Puzzle[51]; Puzzle[51] = Puzzle[52]; Puzzle[52] = Puzzle[53]; Puzzle[53] = temp;
   temp = Puzzle[60]; Puzzle[60] = Puzzle[61]; Puzzle[61] = Puzzle[62]; Puzzle[62] = temp;
   temp = Puzzle[69]; Puzzle[69] = Puzzle[70]; Puzzle[70] = Puzzle[71]; Puzzle[71] = temp;
   temp = Puzzle[78]; Puzzle[78] = Puzzle[79]; Puzzle[79] = Puzzle[80]; Puzzle[80] = temp;
}

void Switch3Columns867( int Puzzle[])  // 678 -> 867  rotate right
{
   register int temp;
   temp = Puzzle[ 8]; Puzzle[ 8] = Puzzle[ 7]; Puzzle[ 7] = Puzzle[ 6]; Puzzle[ 6] = temp;
   temp = Puzzle[17]; Puzzle[17] = Puzzle[16]; Puzzle[16] = Puzzle[15]; Puzzle[15] = temp;
   temp = Puzzle[26]; Puzzle[26] = Puzzle[25]; Puzzle[25] = Puzzle[24]; Puzzle[24] = temp;
   temp = Puzzle[35]; Puzzle[35] = Puzzle[34]; Puzzle[34] = Puzzle[33]; Puzzle[33] = temp;
   temp = Puzzle[44]; Puzzle[44] = Puzzle[43]; Puzzle[43] = Puzzle[42]; Puzzle[42] = temp;
   temp = Puzzle[53]; Puzzle[53] = Puzzle[52]; Puzzle[52] = Puzzle[51]; Puzzle[51] = temp;
   temp = Puzzle[62]; Puzzle[62] = Puzzle[61]; Puzzle[61] = Puzzle[60]; Puzzle[60] = temp;
   temp = Puzzle[71]; Puzzle[71] = Puzzle[70]; Puzzle[70] = Puzzle[69]; Puzzle[69] = temp;
   temp = Puzzle[80]; Puzzle[80] = Puzzle[79]; Puzzle[79] = Puzzle[78]; Puzzle[78] = temp;
}

