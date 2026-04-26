//! F1 help system. Topic content is transcribed from the original
//! Lotus 1-2-3 Release 3.4a `123.HLP` Help screens captured via DOSBox-X
//! (see `tests/help_groundtruth/`). Cross-reference IDs replicate the
//! topic graph encoded in those screens; "index" is the central hub.
//!
//! Rendering and key handling live in [`crate::app`]; this module
//! holds only the topic data.
//!
//! Body text uses `\n` for line breaks. Bodies are kept short enough
//! that the overlay can show them without paging on a default 80×25
//! screen, but the renderer handles vertical scroll for completeness.
//!
//! Topic ordering matches the R3.4a Help Index reading order
//! (column-major: column 1, column 2, column 3). `Tab` walks this slice.
#[derive(Debug, Clone, Copy)]
pub struct HelpTopic {
    pub id: &'static str,
    pub title: &'static str,
    pub body: &'static str,
    /// IDs of related topics ("See also"), as captured from the
    /// authentic R3.4a help-screen footers and inline links.
    pub cross_refs: &'static [&'static str],
}

/// Static topic table. Tab/Shift-Tab walks this slice in order.
pub const HELP_TOPICS: &[HelpTopic] = &[
    HelpTopic {
        id: "index",
        title: "1-2-3 Help Index",
        body: "\
1-2-3 Help Index

About 1-2-3 Help        Macro Basics            1-2-3 Commands
Add-In Commands         Macro Command Index     /Copy
Background Printing     Macro Key Names         /Data
Cell References         Mode Indicators         /File
Control Panel           Printer Information     /Graph
Data Entry              Range Basics            /Move
Data Protection         Recalculation           /Print
Error Message Index     Record Feature          /Quit
File Linking            Status Indicators       /Range
Formulas                Task Index              /System
@Function Index         Undo Feature            /Worksheet
Function Keys           Using Wysiwyg
Keyboard Index

To select a topic, press a pointer-movement key to highlight the topic
and press ENTER. To return to a previous Help screen, press BACKSPACE.
To leave Help and return to the worksheet, press ESC.
",
        cross_refs: &[
            "about-help",
            "add-ins",
            "background-printing",
            "cell-references",
            "control-panel",
            "data-entry",
            "data-protection",
            "error-messages",
            "file-linking",
            "formulas",
            "function-index",
            "function-keys",
            "keyboard-index",
            "macro-basics",
            "macro-commands",
            "macro-keys",
            "mode-indicators",
            "printer-info",
            "range-basics",
            "recalculation",
            "record-feature",
            "status-indicators",
            "task-index",
            "undo-feature",
            "wysiwyg",
            "commands",
            "copy",
            "data",
            "file",
            "graph",
            "move",
            "print",
            "quit",
            "range",
            "system",
            "worksheet",
        ],
    },
    HelpTopic {
        id: "about-help",
        title: "About 1-2-3 Help",
        body: "\
About 1-2-3 Help -- You can view Help screens any time during a 1-2-3
session by pressing F1 (HELP). The 1-2-3 Help system is context-sensitive,
which means that when you press F1 (HELP), 1-2-3 displays a screen that
contains information related to what you are doing in the program.

Help screens contain text and topics in bold. The topics are
cross-references to other screens. To select a topic, highlight the
topic and press ENTER.

Use the following keys to navigate through Help:

  up/down/left/right  Highlight the next topic in that direction.
  BACKSPACE           Displays the previous Help screen.
  END                 Highlights the last Help topic on the screen.
  ENTER               Displays the screen for the highlighted topic.
  ESC                 Leaves Help and returns you to the worksheet.
  F1 (HELP)           Displays the first screen you saw when you pressed F1.
  HOME                Highlights the first Help topic on the screen.

Help is also context-sensitive for 1-2-3 commands, @functions, macro
commands, and error messages. Press F1 (HELP) in any of those contexts
to jump to the matching Help screen.
",
        cross_refs: &["index"],
    },
    HelpTopic {
        id: "add-ins",
        title: "Add-In Commands",
        body: "\
Add-In Commands -- Press ALT-F10 (ADDIN) to display the add-in commands.

  Load     Reads an add-in file into memory.
  Remove   Deletes a specified add-in file from memory.
  Invoke   Starts an add-in application.
  Table    Lists add-in applications, @functions, or macro commands.
  Clear    Removes all add-in files from memory.
  Settings Specifies add-in files 1-2-3 reads automatically and the
           default add-in directory.
  Quit     Leaves the ALT-F10 (ADDIN) menu and returns 1-2-3 to READY.

When you select a command for which you need to specify a file name,
1-2-3 displays a list of files in the current or default add-in
directory.
",
        cross_refs: &["error-messages", "index"],
    },
    HelpTopic {
        id: "background-printing",
        title: "Background Printing",
        body: "\
/Print Background -- Creates an encoded file and prints it while you
continue to work in the worksheet or create additional print jobs.
1-2-3 deletes this encoded file after it finishes printing.

NOTE  You cannot use /Print Background to print on a network printer.

  1. If you did not start BPrint before 1-2-3, save your work, /Quit,
     enter `bprint` at the DOS prompt, and restart 1-2-3.
  2. Select /Print Background.
  3. Specify a file name.
  4. Select Range and specify a range to print.
  5. (Optional) Select Options for different settings.
  6. Make sure the printer is online and at the top of a page.
  7. Select Align, then Go to print, then Quit.
",
        cross_refs: &["print", "index"],
    },
    HelpTopic {
        id: "cell-references",
        title: "Cell References",
        body: "\
Types of Cell and Range References -- To create mixed or absolute
references in a formula, either type a $ (dollar sign) in the cell
reference or press F4 (ABS) when entering or editing a formula.

  Relative reference  Identifies a cell by its position relative to the
                      current cell. A relative reference, like A:B2,
                      changes when you copy a formula that contains it.
  Absolute reference  Contains a $ in front of the worksheet letter,
                      column letter, and row number, e.g. $A:$B$2. An
                      absolute reference always refers to the same cell
                      when you copy a formula that contains it.
  Mixed reference     Contains a $ in front of one or two of the parts.
                      The $-prefixed parts stay the same and the rest
                      change when you copy the formula.
",
        cross_refs: &["formulas", "index"],
    },
    HelpTopic {
        id: "control-panel",
        title: "Control Panel",
        body: "\
The Control Panel consists of the top three lines of the screen.

Line 1 displays information about the current cell and the 1-2-3 mode
indicator. Example layout:

  A:B1:(C2) U [W15] 6500                                          READY

  - Cell address (worksheet letter, column, row)
  - Cell format (if /Range Format was used)
  - Protection status (U = unprotected)
  - Column width (if /Worksheet Column was used)
  - Cell contents
  - Mode indicator (READY, LABEL, VALUE, etc.)

Line 2 displays the characters you are typing or editing, the 1-2-3
menu (after pressing /), or a prompt for command input.

Line 3 displays a submenu, a description of the highlighted command,
or a list of names if the command prompts for a file/range/graph.
",
        cross_refs: &["commands", "index"],
    },
    HelpTopic {
        id: "data-entry",
        title: "Data Entry",
        body: "\
Data Entry -- When 1-2-3 is in READY mode and you start typing, 1-2-3
classifies the entry as either a label or a value based on the first
character. The mode indicator changes to LABEL or VALUE accordingly.

Text entries are labels. Number and formula entries are values.

To enter data:
  1. Move the cell pointer to the cell.
  2. Type the data, up to 512 characters.
  3. Press ENTER or any pointer-movement key to commit.

If 1-2-3 beeps when you press ENTER, you probably made an error. Edit
your entry or press ESC to clear and start over.
",
        cross_refs: &["index"],
    },
    HelpTopic {
        id: "data-protection",
        title: "Data Protection",
        body: "\
Protecting Data and Files

/File Admin Seal       Seals worksheet and reservation settings in a
                       file. You can read a sealed file but cannot
                       change the sealed settings.
File Password          Use /File Save and /File Xtract to save with a
                       password.
File Reservation       Prevents more than one user from saving a shared
                       file at the same time.
/Worksheet Global Prot Globally protects worksheet data so the file
                       can be read but protected cells cannot be
                       changed.
",
        cross_refs: &["file", "index"],
    },
    HelpTopic {
        id: "error-messages",
        title: "Error Message Index",
        body: "\
Error Message Index (excerpt)

  All modified files must be reserved before saving
  Ambiguous field reference in query
  A range name cannot begin with <<>>
  At least one variable range must be specified
  Background printing is currently active
  Backup file exists in memory
  Backup unsuccessful due to file error
  Bin range cannot span worksheets
  BPrint is not in memory
  Break
  Cannot convert extension. File already exists in memory.
  Cannot create file
  Cannot create names from a range of labels not in memory
  Cannot delete all visible worksheets
  Cannot delete all worksheets
  Cannot erase file. It is in use or has read-only access.
  Cannot execute query with a circular reference
  Cannot have more than 256 worksheets in memory
  Cannot read Help file
",
        cross_refs: &["index"],
    },
    HelpTopic {
        id: "file-linking",
        title: "File Linking",
        body: "\
Linking Files with Formulas -- You create a link between two files
when you enter a formula in one file that refers to a cell or range in
another file.

To link, type a file reference in front of the cell or range reference.
A file reference is the name and extension of the file, with or without
a path, enclosed in <<>> (double angle brackets). For example:

  @SUM(<<C:\\SALES\\EAST.WK3>>TOTALS)

To ensure the link works even if you change directories, include the
full path. To share files with others, omit the path so the link
works as long as the linked files are in the current directory.

NOTE  Formulas linked to other files do not automatically update.
Use /File Admin Link-Refresh after /File Open or /File Retrieve to
recalculate them.
",
        cross_refs: &["formulas", "index"],
    },
    HelpTopic {
        id: "formulas",
        title: "Formulas",
        body: "\
Types of Formulas -- 1-2-3 lets you enter four types of formulas:

  Numeric    Performs calculations with numbers using arithmetic
             operators. Example: +B5*5
  String     Performs calculations on text. Example: +\"Mr. \"&B2
  Logical    Performs true/false tests; returns 1 if true, 0 if false.
             Example: +A1>500
  @Function  Performs database, date-and-time, financial, logical,
             mathematical, statistical, scientific, or string
             calculations. Example: @SUM(B10..F10)
",
        cross_refs: &["function-index", "index"],
    },
    HelpTopic {
        id: "function-index",
        title: "@Function Index",
        body: "\
@Function Index (excerpt)

  @@           @COS         @DSTDS       @INDEX       @MAX     @SIN
  @ABS         @COUNT       @DSUM        @INFO        @MID     @SLN
  @ACOS        @CTERM       @D360        @INT         @MIN     @SQRT
  @ASIN        @DATE        @DVAR        @IRR         @MOD     @STD
  @ATAN        @DATEVALUE   @DVARS       @ISERR       @NA      @SUM
  @AVG         @DAVG        @ERR         @ISNA        @NOW     @TAN
  @CELL        @DAY         @EXACT       @ISNUMBER    @NPV     @TODAY
  @CELLPOINTER @DCOUNT      @EXP         @ISRANGE     @PI      @TRIM
  @CHAR        @DDB         @FALSE       @ISSTRING    @PMT     @TRUE
  @CHOOSE      @DGET        @FIND        @LEFT        @PV      @VALUE
  @CODE        @DMAX        @FV          @LENGTH      @RAND    @VAR
  @COLS        @DMIN        @HLOOKUP     @LN          @RATE    @VLOOKUP
  @COORD       @DSTD        @IF          @LOG         @ROUND   @YEAR

Categories: Database, Date-and-Time, Financial, Logical, Mathematical,
Special, Statistical, String.
",
        cross_refs: &["formulas", "index"],
    },
    HelpTopic {
        id: "function-keys",
        title: "Function Keys",
        body: "\
Function Keys -- Perform special operations.

  ALT-F1 (COMPOSE)  Creates characters not on the keyboard.
  ALT-F2 (RECORD)   Replays the record buffer; toggles STEP mode.
  ALT-F3 (RUN)      Selects a macro to run.
  ALT-F4 (UNDO)     Cancels changes since 1-2-3 was last in READY.
  ALT-F6 (ZOOM)     Toggles full-screen view of the current window.
  ALT-F7..F9 (APP)  Start an available add-in.
  ALT-F10 (ADDIN)   Displays the add-in menu.
  CTRL-F9/F10       Display / activate the icon palette (Wysiwyg).

  F1 (HELP)         Displays a Help screen.
  F2 (EDIT)         Switches to EDIT mode for the current entry.
  F3 (NAME)         Lists names of files, graphs, ranges, @functions,
                    macro key names, macro commands, or settings sheets.
  F4 (ABS)          Cycles formula references: relative -> absolute -> mixed.
  F5 (GOTO)         Moves the cell pointer to a cell, worksheet, or file.
  F6 (WINDOW)       Moves the cell pointer between windows.
  F7 (QUERY)        Repeats the last /Data Query command.
  F8 (TABLE)        Repeats the last /Data Table command.
  F9 (CALC)         Recalculates formulas (READY); converts formula to
                    its value (EDIT/VALUE).
  F10 (GRAPH)       Displays the current graph or creates one.
",
        cross_refs: &["keyboard-index", "index"],
    },
    HelpTopic {
        id: "keyboard-index",
        title: "Keyboard Index",
        body: "\
Keyboard Index

  Change entries                     -- see Editing Keys
  Move around the worksheet          -- see Pointer-Movement Keys
  Move between worksheets in a file  -- see Worksheet Pointer-Movement Keys
  Move between active files          -- see File Pointer-Movement Keys
  Use F1-F10, ALT, and CTRL          -- see Function Keys
  Use ALT with function keys / run a -- see Special Keys
    macro
  Use CTRL-BREAK and ESC to cancel   -- see Special Keys
",
        cross_refs: &["function-keys", "index"],
    },
    HelpTopic {
        id: "macro-basics",
        title: "Macro Basics",
        body: "\
Macro Basics

A macro is a series of 1-2-3 commands and keystrokes that defines a
1-2-3 task. You enter the macro as one or more labels in a column and
assign it a range name. When you run the macro, 1-2-3 automatically
performs the defined task.

Macros save time by performing simple but repetitive tasks
automatically. You can write macros to automate complex procedures and
guide users who are unfamiliar with 1-2-3 through specific applications.

  Creating a Macro    How to choose a location for, enter, name,
                      document, run, debug, and save a macro.
  The Record Feature  How to use ALT-F2 (RECORD) to write a macro and
                      play back the keystrokes for a task.
  Macro Index         List of macro commands and descriptions.
",
        cross_refs: &["record-feature", "macro-commands", "index"],
    },
    HelpTopic {
        id: "macro-commands",
        title: "Macro Command Index",
        body: "\
Macro Command Index

  subroutine   DISPATCH    GRAPHON     READ
  ?            FILESIZE    IF          READLN
  APPENDBELOW  FOR         INDICATE    RECALC
  APPENDRIGHT  FORBREAK    LET         RECALCCOL
  BEEP         FORM        LOOK        RESTART
  BLANK        FORMBREAK   MENUBRANCH  RETURN
  BRANCH       FRAMEOFF    MENUCALL    SETPOS
  BREAK        FRAMEON     ONERROR     SYSTEM
  BREAKOFF     GET         OPEN        WAIT
  BREAKON      GETLABEL    PANELOFF    WINDOWSOFF
  CLOSE        GETNUMBER   PANELON     WINDOWSON
  CONTENTS     GETPOS      PUT         WRITE
  DEFINE       GRAPHOFF    QUIT        WRITELN
                                       /X commands

NOTE  Optional macro arguments are enclosed in [ ] (brackets). All
other arguments are required.
",
        cross_refs: &["macro-basics", "macro-keys", "index"],
    },
    HelpTopic {
        id: "macro-keys",
        title: "Macro Key Names",
        body: "\
Macro Key Names (excerpt)

  1-2-3 key                  Macro key name
  ----------------------     -------------------------------
  down arrow                 {DOWN} or {D}
  up arrow                   {UP} or {U}
  left arrow                 {LEFT} or {L}
  right arrow                {RIGHT} or {R}
  ALT-F6 (ZOOM)              {ZOOM}
  ALT-F10 (ADDIN)            {ADDIN} or {APP4}
  BACKSPACE                  {BACKSPACE} or {BS}
  CTRL-left or SHIFT-TAB     {BIGLEFT}
  CTRL-right or TAB          {BIGRIGHT}
  CTRL-END                   {FILE}
  CTRL-END CTRL-PG DN        {PREVFILE}, {PF}, or {FILE}{PS}
  CTRL-END CTRL-PG UP        {NEXTFILE}, {NF}, or {FILE}{NS}
  CTRL-HOME                  {FIRSTCELL} or {FC}
  DEL                        {DELETE} or {DEL}
  END                        {END}
  ESC                        {ESCAPE} or {ESC}
  F1 (HELP)                  {HELP}
  F2 (EDIT)                  {EDIT}
",
        cross_refs: &["macro-basics", "macro-commands", "index"],
    },
    HelpTopic {
        id: "mode-indicators",
        title: "Mode Indicators",
        body: "\
Mode Indicator -- 1-2-3 displays the mode indicator in the upper right
corner of the screen.

  EDIT    F2 (EDIT) was pressed, /Data Parse Format-Line Edit was
          selected, or you made an incorrect entry.
  ERROR   1-2-3 is displaying an error message. F1 for Help; ESC or
          ENTER to clear.
  FILES   1-2-3 is displaying a list of file names.
  FIND    /Data Query Find or F7 (QUERY).
  HELP    F1 (HELP) was pressed.
  LABEL   You are entering a label.
  MENU    1-2-3 is displaying a menu of commands (after /).
  NAMES   1-2-3 is displaying a list of names.
  POINT   You are highlighting a range.
  READY   Ready for data entry or a command.
  STAT    A status screen is displayed.
  VALUE   You are entering a value.
  WAIT    1-2-3 is completing a command (e.g. saving a file).
  WYSIWYG : (colon) was pressed; Wysiwyg menu is open.
",
        cross_refs: &["error-messages", "index"],
    },
    HelpTopic {
        id: "printer-info",
        title: "Printer Information",
        body: "\
Printer Information -- The number of lines 1-2-3 prints per page varies
by printer. Set page length with /Print [B,E,F,P] Options Pg-Length.

  Printer                       Portrait      Landscape
  Apple LaserWriter                63          47
  Dot-Matrix Printers              66          Wysiwyg only
  HP DeskJet                       60          45
  HP LaserJet (no font cart.)      60          45
  HP LaserJet (F/J/Z cart.)        72          45
  HP PaintJet (continuous)         66          Wysiwyg only
  HP PaintJet (single feed)        60          Wysiwyg only
  HP Plotters 8.5x11               60          45
  HP Plotters 11x17                95          60

NOTE  You can use Wysiwyg (:Print) to print in landscape mode on most
printers.
",
        cross_refs: &["print", "index"],
    },
    HelpTopic {
        id: "range-basics",
        title: "Range Basics",
        body: "\
Range Basics

A range is any rectangular block of adjacent cells. A range can be a
single cell (A1..A1), part of a row (A2..C2) or column (B1..B3), or
a block of cells (A1..C3). A range can be three-dimensional, spanning
adjacent worksheets (A:A1..C:C3) as long as the worksheets are in the
same file.

A range address consists of any two diagonally opposite corners,
separated by one or two periods. For example, A1..A1 and A:A1..A:C3
are valid range addresses.
",
        cross_refs: &["range", "index"],
    },
    HelpTopic {
        id: "recalculation",
        title: "Recalculation",
        body: "\
/Worksheet Global Recalc -- Controls when, in what order, and how many
passes 1-2-3 uses when recalculating formulas.

Order:
  Natural     (Default) Recalculates dependencies first.
  Columnwise  Column by column starting at A:A1.
  Rowwise     Row by row starting at A:A1.

Mode:
  Automatic   (Default) Each change recalculates affected formulas in
              the background.
  Manual      Recalculates only when you press F9 (CALC).
  Iteration   Sets passes 1-50 (used with Columnwise/Rowwise, or
              Natural with a circular reference).

Settings affect all active files. Use /File Admin Link-Refresh to
update formulas linked to files on disk.
",
        cross_refs: &["worksheet", "index"],
    },
    HelpTopic {
        id: "record-feature",
        title: "Record Feature",
        body: "\
The Record Feature -- ALT-F2 (RECORD) gives you access to the record
buffer, a 512-byte area that records your keystrokes. Press ALT-F2
(RECORD) and select one of:

  Playback  Repeats the keystrokes in the record buffer.
  Copy      Copies keystrokes to the worksheet (auto-write a macro).
  Erase     Clears the record buffer.
  Step      Runs a macro one step at a time (debugging).
  Trace     Runs a macro and shows which step failed.
",
        cross_refs: &["macro-basics", "index"],
    },
    HelpTopic {
        id: "status-indicators",
        title: "Status Indicators",
        body: "\
Status Indicators -- Displayed in the status line at the bottom of the
screen.

  CALC    Recalculation needed (blue) or in progress (red).
  CAP     CAPS LOCK is on.
  CIRC    A formula contains a circular reference.
  CMD     A macro is running.
  END     END key was pressed; awaiting a pointer-movement key.
  FILE    CTRL-END is moving between files.
  GROUP   The current file is in GROUP mode.
  MEM     Available memory has fallen below 32 KB or is fragmented.
  NUM     NUM LOCK is on.
  OVR     Overstrike (insert) mode (toggled by INS).
  PRT     Background printing is in progress.
  RO      The current file has read-only status.
  SCROLL  SCROLL LOCK is on.
  STEP    A macro is running in STEP mode.
  ZOOM    A window is zoomed full-screen via ALT-F6.

The file-and-clock indicator appears in the lower left. Set its display
with /Worksheet Global Default Other Clock.
",
        cross_refs: &["index"],
    },
    HelpTopic {
        id: "task-index",
        title: "Task Index",
        body: "\
Task Index (excerpt)

  Change column width                  -- /Worksheet Column
  Change display of data               -- /Range Format
  Change formulas to numbers           -- /Range Value
  Change global default settings       -- /Worksheet Global Default
  Create a multiple-sheet file         -- /Worksheet Insert
  Create a text file                   -- /Print File
  Create a what-if table               -- /Data Table
  Display three worksheets at once     -- /Worksheet Window
  End a 1-2-3 session                  -- /Quit
  Enter dates                          -- Date Formats
  Enter formulas                       -- Entering Formulas
  Enter information in a worksheet     -- Data Entry
  Erase a file on disk                 -- /File Erase
  Erase worksheet data                 -- /Range Erase
  Find a circular reference            -- /Worksheet Status
  Find records in a database           -- /Data Query Find
  Get Help with 1-2-3                  -- About Help
  Hide data                            -- Hiding Data
",
        cross_refs: &["index"],
    },
    HelpTopic {
        id: "undo-feature",
        title: "Undo Feature",
        body: "\
Using Undo -- When undo is on, ALT-F4 (UNDO) cancels the effects of
the most recent operation that changed worksheet data or settings.
ALT-F4 (UNDO) is not a toggle key; using it twice does not restore.

You can usually undo even after starting another entry by escaping
out with ESC until 1-2-3 returns to READY mode and then pressing
ALT-F4.

NOTE  Undo cancels changes to active worksheets only. It cannot undo
file activity, printer activity, cell-pointer movement, or the
recalculation triggered by F9 (CALC) or /File Admin Link-Refresh.

ALT-F4 (UNDO) works from the keyboard only, not from a macro.
",
        cross_refs: &["worksheet", "index"],
    },
    HelpTopic {
        id: "wysiwyg",
        title: "Using Wysiwyg",
        body: "\
Using Wysiwyg -- Wysiwyg is a spreadsheet publishing add-in that
enhances the appearance of worksheets. Wysiwyg includes SmartIcons,
which provide easy access to 1-2-3 and Wysiwyg commands.

To display the Wysiwyg main menu, press : (colon).

To display Help for the Wysiwyg main menu, press : (colon) and then
F1 (HELP).

NOTE  If Wysiwyg is not in memory, pressing : does not display the
Wysiwyg menu. Use ALT-F10 (ADDIN) Load to bring it back, or set
automatic start-up via ALT-F10 (ADDIN) Settings System.
",
        cross_refs: &["add-ins", "index"],
    },
    HelpTopic {
        id: "commands",
        title: "1-2-3 Commands",
        body: "\
1-2-3 Commands

  Worksheet commands       Print commands
  Range commands           Graph commands
  Copy command             Data commands
  Move command             System command
  File commands            Quit command

To use 1-2-3 commands, press / (slash) to display the main menu in the
control panel. Highlight a command and press ENTER, or press the
first character. The third line of the control panel shows an
explanation or submenu for the highlighted command.

To back out of a menu one level, press ESC.
To leave a menu and return 1-2-3 to READY mode, press CTRL-BREAK.
",
        cross_refs: &[
            "worksheet",
            "range",
            "copy",
            "move",
            "file",
            "print",
            "graph",
            "data",
            "system",
            "quit",
            "index",
        ],
    },
    HelpTopic {
        id: "copy",
        title: "/Copy",
        body: "\
/Copy -- Copies a range of data, including cell formats and protection
status, to a range in the same file or in a different file.

  1. (Optional) Move the pointer to the first cell to copy.
  2. Select /Copy.
  3. Specify the range to copy FROM.
  4. Specify the range to copy TO.

To make one copy, specify only one cell as the TO range. To make
multiple copies, specify a range as the TO range.

CAUTION  1-2-3 replaces existing data in the TO range without warning.

When copying formulas, 1-2-3 adjusts cell references depending on
their kind (absolute, relative, or mixed). Within the same file,
relative references shift; absolute references do not. Across files,
all reference parts adjust to the new file.
",
        cross_refs: &["cell-references", "range-basics", "index"],
    },
    HelpTopic {
        id: "data",
        title: "/Data",
        body: "\
The Data Commands -- Analyze and manipulate data in worksheets, 1-2-3
database tables, and external databases.

  Fill         Fills a range with a sequence of values.
  Table        Creates a what-if table from formulas.
  Sort         Arranges records in a database table.
  Query        Locates and edits selected records.
  Distribution Creates a frequency distribution.
  Matrix       Inverts or multiplies matrixes.
  Regression   Multiple linear regression analysis (up to 75 vars).
  Parse        Splits a column of long labels into several columns.
  External     Connects 1-2-3 and external database tables.
",
        cross_refs: &["index"],
    },
    HelpTopic {
        id: "file",
        title: "/File",
        body: "\
The File Commands

  Retrieve  Reads a worksheet file into memory, replacing the current
            file.
  Save      Saves worksheet files on disk.
  Combine   Incorporates data from a file on disk into the current file.
  Xtract    Extracts a range of data and saves it in a worksheet file.
  Erase     Erases a file on disk.
  List      Displays a temporary list of file information.
  Import    Reads data from a text file into the current worksheet.
  Dir       Changes the directory for reads/saves/lists.
  New       Creates a new blank worksheet file.
  Open      Reads a file into memory alongside others.
  Admin     Reservations, file table, sealing, link refresh.
",
        cross_refs: &["data-protection", "index"],
    },
    HelpTopic {
        id: "graph",
        title: "/Graph",
        body: "\
The Graph Commands -- Define, display, and save 1-2-3 graphs.

  Type     Specifies the graph type and orientation.
  X        Specifies the x-axis labels, values, or pie-slice labels.
  A - F    Specifies the numeric data ranges to graph.
  Reset    Resets some or all current graph settings.
  View     Draws a full-screen view of the current graph.
  Save     Saves the current graph for use with other programs.
  Options  Adds enhancements and sets axis scaling.
  Name     Creates, retrieves, deletes, and lists named graphs.
  Group    Assigns all data ranges (X and A-F) at once.
  Quit     Returns 1-2-3 to READY mode.
",
        cross_refs: &["print", "index"],
    },
    HelpTopic {
        id: "move",
        title: "/Move",
        body: "\
/Move -- Transfers a range of data, including cell formats and
protection status, to another range in the same file. /Move cannot
transfer data between files.

  1. (Optional) Move the pointer to the first cell of the FROM range.
  2. Select /Move.
  3. Specify the range to move FROM.
  4. Specify the range to move TO. Only the first cell is needed.

CAUTION  1-2-3 replaces existing data in the TO range. Moving data
into or out of the first or last cell of a named or referenced range
changes the range definition.
",
        cross_refs: &["range-basics", "index"],
    },
    HelpTopic {
        id: "print",
        title: "/Print",
        body: "\
The Print Commands

  Printer    Selects a printer as the destination.
  File       Selects a text file as the destination.
  Encoded    Selects an encoded file as the destination.
  Background Selects background printing as the destination.
  Suspend    Temporarily stops background printing.
  Resume     Continues a suspended print job and clears errors.
  Cancel     Cancels all 1-2-3 background printing.
  Quit       Returns 1-2-3 to READY mode.
",
        cross_refs: &["printer-info", "index"],
    },
    HelpTopic {
        id: "quit",
        title: "/Quit",
        body: "\
/Quit -- Ends the current 1-2-3 session and returns you to DOS.

CAUTION  /Quit cancels all pending print jobs and clears worksheets
from memory without saving. Use /File Save first if you want to
preserve your work.

  1. Select /Quit.
  2. Select No to return to READY mode, or Yes to end 1-2-3.

If unsaved changes or pending print jobs exist, 1-2-3 displays
another No/Yes confirmation:
  - No cancels /Quit so you can save the worksheets.
  - Yes ends 1-2-3 without saving.
",
        cross_refs: &["file", "index"],
    },
    HelpTopic {
        id: "range",
        title: "/Range",
        body: "\
The Range Commands

  Format   Changes the display of data in a range.
  Label    Left-aligns, right-aligns, or centers labels.
  Erase    Erases data in a range.
  Name     Creates, deletes, undefines range names; manages notes.
  Justify  Rearranges a column of labels to fit a specified width.
  Prot     Prevents changes to unprotected cells when global protection
           is on.
  Unprot   Allows changes to cells when global protection is on.
  Input    Restricts pointer movement to unprotected cells.
  Value    Copies a range, converting formulas to current values.
  Trans    Copies a range, transposing rows and columns.
  Search   Finds or replaces strings in a range.
",
        cross_refs: &["range-basics", "index"],
    },
    HelpTopic {
        id: "system",
        title: "/System",
        body: "\
/System -- Temporarily suspends 1-2-3 and returns you to DOS so you
can use DOS commands without ending the current 1-2-3 session.

  1. Use /File Save first if you want to save your work.
  2. Select /System. 1-2-3 replaces the worksheet with the DOS prompt.
     You can copy files, create directories, run programs, etc.

CAUTION  Do not load memory-resident programs while in /System; you
may not be able to resume the 1-2-3 session.

  3. Type `exit` and press ENTER at the DOS prompt to return to 1-2-3.
",
        cross_refs: &["index"],
    },
    HelpTopic {
        id: "worksheet",
        title: "/Worksheet",
        body: "\
The Worksheet Commands

  Global   Establishes global settings for worksheets and files;
           changes 1-2-3 configuration settings.
  Insert   Inserts blank columns, rows, or worksheets.
  Delete   Deletes columns, rows, worksheets, and files from memory.
  Column   Sets / resets column widths; hides / redisplays columns.
  Erase    Erases all active worksheets and files from memory.
  Titles   Freezes rows and columns at the top and/or left edges.
  Window   Provides different ways to view worksheets.
  Status   Memory, hardware, recalculation, global settings.
  Page    Creates page breaks in printed worksheets.
  Hide    Hides and redisplays worksheets.
",
        cross_refs: &["recalculation", "index"],
    },
];

/// Lookup a topic by id. Returns None if the id is unknown.
pub fn find_topic(id: &str) -> Option<&'static HelpTopic> {
    HELP_TOPICS.iter().find(|t| t.id == id)
}

/// All topics available to the F1 overlay: the curated 32-entry slice
/// followed by every decoded entry whose title isn't already covered
/// by a curated topic. The merged slice gives the user navigation
/// over every topic the original `123.HLP` knows about (~321 total),
/// while keeping the curated bodies for the high-level topics where
/// they exist.
///
/// Decoded entries get synthetic ids of the form `decoded-<slug>`,
/// no cross-refs (the cross-reference graph for decoded topics isn't
/// fully recovered yet), and bodies that may contain residual
/// renderer-control noise — see `docs/HLP_DECODE_NOTES.md`.
pub fn all_topics() -> &'static [HelpTopic] {
    use std::sync::OnceLock;
    static FULL: OnceLock<Vec<HelpTopic>> = OnceLock::new();
    FULL.get_or_init(|| {
        let mut out: Vec<HelpTopic> = HELP_TOPICS.to_vec();
        let curated_titles: std::collections::HashSet<&str> =
            HELP_TOPICS.iter().map(|t| t.title).collect();
        let mut used_ids: std::collections::HashSet<String> =
            HELP_TOPICS.iter().map(|t| t.id.to_string()).collect();
        for &(title, body) in crate::help_topics_decoded::HELP_TOPICS_DECODED {
            if curated_titles.contains(title) {
                continue;
            }
            // Synthesize a unique slug-based id.
            let mut id = format!("decoded-{}", slugify(title));
            let mut suffix = 2;
            while used_ids.contains(&id) {
                id = format!("decoded-{}-{}", slugify(title), suffix);
                suffix += 1;
            }
            used_ids.insert(id.clone());
            out.push(HelpTopic {
                id: Box::leak(id.into_boxed_str()),
                title,
                body,
                cross_refs: &[],
            });
        }
        out
    })
}

/// Lower-case, hyphen-separated slug for synthetic topic ids.
fn slugify(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut prev_dash = false;
    for c in s.chars() {
        if c.is_ascii_alphanumeric() {
            out.push(c.to_ascii_lowercase());
            prev_dash = false;
        } else if !prev_dash && !out.is_empty() {
            out.push('-');
            prev_dash = true;
        }
    }
    while out.ends_with('-') {
        out.pop();
    }
    if out.is_empty() {
        out.push_str("untitled");
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn topic_table_nonempty_and_unique_ids() {
        assert!(!HELP_TOPICS.is_empty());
        let mut ids: Vec<&str> = HELP_TOPICS.iter().map(|t| t.id).collect();
        ids.sort();
        let count = ids.len();
        ids.dedup();
        assert_eq!(ids.len(), count, "duplicate topic ids");
    }

    #[test]
    fn first_topic_is_index() {
        assert_eq!(HELP_TOPICS[0].id, "index");
    }

    #[test]
    fn find_topic_lookup() {
        assert!(find_topic("index").is_some());
        assert!(find_topic("function-keys").is_some());
        assert!(find_topic("nope").is_none());
    }

    #[test]
    fn every_topic_has_nonempty_title_and_body() {
        for t in HELP_TOPICS {
            assert!(!t.title.is_empty(), "title empty: {}", t.id);
            assert!(!t.body.is_empty(), "body empty: {}", t.id);
        }
    }

    #[test]
    fn topics_have_cross_references() {
        for t in HELP_TOPICS {
            if t.id == "index" {
                assert!(
                    t.cross_refs.len() >= 5,
                    "index should link to many topics, got {}",
                    t.cross_refs.len()
                );
            } else {
                assert!(
                    t.cross_refs.contains(&"index"),
                    "topic {} should link back to index",
                    t.id
                );
            }
            for xref in t.cross_refs {
                assert!(
                    HELP_TOPICS.iter().any(|p| p.id == *xref),
                    "topic {} references unknown id {}",
                    t.id,
                    xref
                );
            }
        }
    }

    #[test]
    fn about_help_has_authentic_text() {
        let t = find_topic("about-help").expect("about-help topic");
        assert!(t.body.contains("About 1-2-3 Help"));
        assert!(t.body.contains("context-sensitive"));
    }

    #[test]
    fn all_topics_merges_curated_and_decoded() {
        let all = all_topics();
        // Should be at least 200 topics (curated 32 + most of decoded 321).
        assert!(all.len() >= 200, "all_topics len {} too small", all.len());
        // First entry remains the curated index page.
        assert_eq!(all[0].id, "index");
        // Curated entries appear as-is at the start.
        for (i, c) in HELP_TOPICS.iter().enumerate() {
            assert_eq!(all[i].id, c.id);
            assert_eq!(all[i].title, c.title);
        }
        // No duplicate ids across the whole merged slice.
        let mut ids: Vec<&str> = all.iter().map(|t| t.id).collect();
        ids.sort();
        let n = ids.len();
        ids.dedup();
        assert_eq!(ids.len(), n, "duplicate ids in all_topics");
        // No duplicate titles either.
        let mut titles: Vec<&str> = all.iter().map(|t| t.title).collect();
        titles.sort();
        let n = titles.len();
        titles.dedup();
        assert_eq!(titles.len(), n, "duplicate titles in all_topics");
    }

    #[test]
    fn all_topics_includes_decoded_only_entries() {
        // Decoded set has many topics not in the curated table —
        // confirm at least a few well-known ones come through.
        let titles: std::collections::HashSet<&str> =
            all_topics().iter().map(|t| t.title).collect();
        for must in &["/File Erase", "File Combine", "File Admin"] {
            assert!(
                titles.contains(must),
                "all_topics missing decoded-only topic {must:?}"
            );
        }
        // The merged slice should be substantially larger than the
        // curated 32 — at least 200 entries from the decoded set.
        assert!(
            all_topics().len() > HELP_TOPICS.len() + 100,
            "merged slice has only {} topics; curated alone has {}",
            all_topics().len(),
            HELP_TOPICS.len()
        );
    }

    #[test]
    fn slugify_examples() {
        assert_eq!(slugify("/Copy"), "copy");
        assert_eq!(slugify("@Function Index"), "function-index");
        assert_eq!(slugify("About 1-2-3 Help"), "about-1-2-3-help");
        assert_eq!(slugify("F3 (NAME) Key"), "f3-name-key");
        assert_eq!(slugify("///"), "untitled");
    }

    #[test]
    fn decoded_topics_table_present_and_nonempty() {
        use crate::help_topics_decoded::HELP_TOPICS_DECODED;
        // The committed decoded set covers every body topic from
        // 123.HLP (~321 entries). It's a fallback / reference source;
        // see docs/HLP_DECODE_NOTES.md for the renderer-noise caveats.
        assert!(
            HELP_TOPICS_DECODED.len() >= 200,
            "expected ≥200 decoded topics, got {}",
            HELP_TOPICS_DECODED.len()
        );
        // Every entry has non-empty title and body.
        for (title, body) in HELP_TOPICS_DECODED {
            assert!(!title.is_empty(), "empty title");
            assert!(!body.is_empty(), "empty body for {title}");
        }
        // Spot-check a few well-known titles are present.
        let titles: std::collections::HashSet<&str> =
            HELP_TOPICS_DECODED.iter().map(|(t, _)| *t).collect();
        for must in &[
            "About 1-2-3 Help",
            "/Copy",
            "Control Panel",
            "Print Commands",
        ] {
            assert!(
                titles.contains(must),
                "decoded set missing expected topic {must:?}"
            );
        }
    }

    #[test]
    fn decoded_topics_carry_authentic_phrases() {
        use crate::help_topics_decoded::HELP_TOPICS_DECODED;
        // Pull the body for "About 1-2-3 Help" out of the committed
        // decoded set and confirm the same distinctive phrase the
        // curated entry carries also appears in the decoded copy.
        let about = HELP_TOPICS_DECODED
            .iter()
            .find(|(t, _)| *t == "About 1-2-3 Help")
            .expect("About 1-2-3 Help in decoded set");
        assert!(
            about.1.starts_with("You can view Help screens any time"),
            "decoded About body started with: {:?}",
            &about.1[..80.min(about.1.len())]
        );
    }

    /// Cross-check: hand-authored topic bodies should agree with the
    /// `123.HLP` decoded source for the same topic. Skipped when the
    /// `.HLP` file isn't present (CI / contributors without the
    /// original DOS install).
    ///
    /// This isn't a strict equality check — the decoder produces text
    /// with renderer-control noise interleaved, while `HELP_TOPICS`
    /// holds the curated transcription. We only assert that distinctive
    /// short phrases from the curated body also appear in the decoded
    /// body, which is enough to confirm provenance.
    #[test]
    fn help_topics_agree_with_decoded_source() {
        let path = std::env::var("L123_HLP_FILE").unwrap_or_else(|_| {
            format!(
                "{}/Documents/dosbox-cdrive/123R34/123.HLP",
                std::env::var("HOME").unwrap_or_default()
            )
        });
        let Ok(bytes) = std::fs::read(&path) else {
            eprintln!("skipping: 123.HLP not found at {path}");
            return;
        };
        let decoded = l123_help::topics(&bytes).expect("decoded topics");

        // (curated_id, distinctive_phrase) — the phrase must appear in
        // BOTH the curated body AND in *some* decoded topic's body.
        // We pick short phrases unlikely to be split by the renderer-
        // control noise that's interleaved into the decoded bytes.
        // We don't try to match titles exactly because the decoder
        // sometimes drops the first character of a title when the
        // renderer encoded it as a soft-space (`0xC4`) — e.g.
        // "Using Undo" becomes "Undo" after decode.
        let pairs: &[(&str, &str)] = &[
            ("about-help", "context-sensitive"),
            ("control-panel", "top three lines"),
            ("copy", "Copies a range of data"),
            ("undo-feature", "ALT-F4"),
        ];

        for (curated_id, phrase) in pairs {
            let curated = find_topic(curated_id)
                .unwrap_or_else(|| panic!("curated topic {curated_id} missing"));
            assert!(
                curated.body.contains(phrase),
                "curated body for {curated_id} missing phrase {phrase:?}"
            );
            let in_decoded = decoded.iter().any(|t| t.body.contains(phrase));
            assert!(
                in_decoded,
                "no decoded topic body contains phrase {phrase:?} (from curated {curated_id})"
            );
        }
    }
}
