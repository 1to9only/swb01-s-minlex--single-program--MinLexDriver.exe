### swb01's minlex - single program - MinLexDriver.exe

This is swb01's minlex program - in a single program: [MinLexDriver.exe](https://github.com/1to9only/swb01-s-minlex--single-program--MinLexDriver.exe/releases).

The program is built from source code (release 4) provided by swb01, with some modifications by me.
- merged MinLex9X9SR1.vb into MinLexDriver.vb, and modified to compile
- fixed (removed) the null character written at end of output file

I only use option 2 - i.e. other options are not used (untested).

Note: The input file size must be a multiple of 83, i.e. each line = 81-chars sudoku + CRLF.

swb01's thread: [http://forum.enjoysudoku.com/minlex-routine-t39261.html](http://forum.enjoysudoku.com/minlex-routine-t39261.html)

swb01's source: [https://drive.google.com/drive/folders/1WpEwW1RUYIly_oCLsmBo-wyU_bvmfpNN?usp=sharing](https://drive.google.com/drive/folders/1WpEwW1RUYIly_oCLsmBo-wyU_bvmfpNN?usp=sharing)

### MinLexDriver.exe - Version: VB-2025_09_09

With modifications by 1to9only

This is shorter and much modified version of the original source code.
It supports old process mode 2 only, no need to specify the option.
Most of unused code, e.g. code for process mode 3, has been removed.

The minlexing code (MinLex9X9SR1) has seen some cosmetic changes.
The program usage has also been changed.

Usage: MinLexDriver.exe inputfile outputfile\
e.g.
```
MinLexDriver.exe
MinLexDriver.exe ......789......1.....2.3...2......65.9..71.............7..9.8..............5...4.
MinLexDriver.exe 17puz49158.txt output.txt

```
For the input file, the restriction of 83-chars for each line remains.

