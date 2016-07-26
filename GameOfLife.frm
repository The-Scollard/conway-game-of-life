VERSION 5.00
Begin VB.Form frmGameOfLife 
   Caption         =   "Game of Life"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   654
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.HScrollBar hsbTimerScroll 
      Height          =   375
      Left            =   6120
      Max             =   2000
      Min             =   16
      TabIndex        =   1
      Top             =   6480
      Value           =   16
      Width           =   3015
   End
   Begin VB.Timer timLife 
      Interval        =   16
      Left            =   5160
      Top             =   6120
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblSlow 
      Caption         =   "Slow"
      Height          =   255
      Left            =   8640
      TabIndex        =   3
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label lblFast 
      Caption         =   "Fast"
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   6960
      Width           =   495
   End
End
Attribute VB_Name = "frmGameOfLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ALife: The Game of Life
'Luke Scollard
'ICS 3U1
'06/24/14

'Purpose: The purpose of this program is to simulate life, with organisms dying from overpopulation and underpopulation,
'and reproducing new organisms. The purpose is also to allow the user to interact with the program, and create patterns
'of life, which they can then see with a visual display.

'Design Decisions:

'A design decision was to allow the user to make cells alive, or kill them, using the mouse and clicking the screen. This
'was chosen because using a WASD and Spacebar setup was harder to implement compared to the mouse, and would be more
'tedious for the user to turn cells on or off than the mouse setup that the program uses.

'Another design decision, albeit one that was required, was the creation of a visual display. The visual display shows the
'user which cells are alive, and which are dead. The visual display was created using the draw functions of Visual Basic,
'another design decision, with dead cells being represented by grey squares, and live cells being represented by red squares.

'Another design decision was to let the user control the intervals where the program calculates which cells are living, and
'which are dead using a horizontal scroll bar. This allows the user to slow down the visual display, if they want to see how
'their life pattern evolves, or speed the visual display up, if they want to run the program quickly.

'Another design decision is to let the user pause the visual display, and change it, before running the program again. This
'allows them to change their display while it is running, allowing them test out different configurations of live and dead
'cells.

'Another design decision is to let the user save their particular pattern of cells into a file, and the ability to load that
'pattern even after the program has been closed. This allows the user, if they have found a particularly interesting
'pattern, to save it for later.

'Another design decision is to have the program follow the exact rules of Conway's Game of Life, with no changes. If a cell
'is alive, and it is surrounded by either 2 or 3 alive cells, it stays alive. If not, it dies. If a dead cell has 3 live
'neighbours, it becomes alive, if not, it stays dead. The reason these rules were chosen was that they are simple, intuitive,
'and easy to understand.

'Variable Dictionary

'Variable Name      Scope       Type        Purpose
'Alive()            General     Boolean     This variable contains the state, alive or dead, of all of the program's 2399
'                                           cells, which each part of the array representing the state of one cell. This
'                                           is used by the program to correctly draw the visual display.
'NextAlive()        General     Boolean     This variable contains the state, alive or dead, each of the program's 2399
'                                           cells will be the next time the alive calculations are made. This is seperate
'                                           from the Alive array so that the state the cells will be next calculation will
'                                           not effect the state of the cell this calculation.
'Started            General     Boolean     This boolean causes the program to start, and calculate which cells are alive,
'                                           when it is true. When the boolean is not true, the program stops, or pauses,
'                                           if the display is running.

Dim Alive(0 To 2399) As Boolean
Dim NextAlive(0 To 2399) As Boolean
Dim Started As Boolean


Private Sub cmdLoad_Click()

'Purpose: The purpose of this button is to load the saved cell pattern when it is clicked by the user.

'Variable Dictionary

'Variable Name      Scope       Type        Purpose
'X                  cmdLoad     Integer     This variable holds the X value, or X position, for the cell being loaded,
'                                           so that the program can find the cell's proper place in the Alive array.
'Y                  cmdLoad     Integer     This variable holds the Y value, or Y position, for the cell being loaded,
'                                           so that the program can find the cell's proper place in the Alive array.
'AliveNumber        cmdLoad     Integer     This variable holds the state of the cell being loaded, whether it is alive
'                                           or dead, with the number 1 representing alive, and the number 0 representing
'                                           dead.
                                            

    Dim X As Integer
    Dim Y As Integer
    Dim AliveNumber As Integer
    Open "U:\Documents\SavedGame.txt" For Input As #2
        Do Until EOF(2) = True
            Input #2, X, Y, AliveNumber
            If AliveNumber = 1 Then         'If the AliveNumber indicates the cell should be alive, the cell is made alive.
                Alive(60 * Y + X) = True
            ElseIf AliveNumber = 0 Then     'If the AliveNumber indicates the cell should be dead, the cell is made dead.
                Alive(60 * Y + X) = False
            End If
        Loop
    Close #2
         
End Sub

Private Sub cmdSave_Click()

'Purpose: The purpose of this button is to save the cell pattern currently displayed when clicked by the user, so it may
'be loaded at a later time.

'Variable Dictionary

'Variable Name      Scope       Type        Purpose
'X                  cmdSave     Integer     This variable is used as a counter in a For Loop, where it holds the X value,
'                                           or X position, to find the cell's position in the Alive array.
'Y                  cmdSave     Integer     This variable is used as a counter in a For Loop, where it holds the Y value,
'                                           or Y position, to find the cell's position in the Alive array.

    Dim X As Integer
    Dim Y As Integer
    Open "U:\Documents\SavedGame.txt" For Output As #1
        For X = 0 To 59
            For Y = 0 To 39
                If Is_Alive(X, Y) = True Then
                    Write #1, X, Y, 1           'Writes X Position, Y Position, Alive
                Else
                    Write #1, X, Y, 0           'Writes X Position, Y Position, Dead
                End If
            Next Y
        Next X
    Close #1
End Sub

Private Sub cmdStart_Click()

'Purpose: The purpose of this button is to start the program's calculations, where it determines the state of all the cells
'in the next run. It also starts the visual display for the program.

    Started = True
    cmdStart.Visible = False
    cmdStop.Visible = True
End Sub

Sub Render()

'Purpose: The purpose of this subroutine is to draw all the components of the visual display, such as the background and the
'alive squares.

'Variable Dictionary

'Variable Name      Scope       Type        Purpose
'X                  Render      Integer     This variable is used as a counter in a For Loop, where it holds the X value,
'                                           or X position, to find the cell's position in the Alive array, so that
'                                           it can be drawn on the visual display if it is alive.
'Y                  Render      Integer     This variable is used as a counter in a For Loop, where it holds the Y value,
'                                           or Y position, to find the cell's position in the Alive array, so that
'                                           it can be drawn on the visual display if it is alive.

    Dim X As Integer
    Dim Y As Integer
    Line (0, 0)-(600, 400), RGB(100, 100, 100), BF      'Draws background
    For X = 0 To 59
        For Y = 0 To 39
            If Alive(60 * Y + X) = True Then        'If cell is alive, the red square for the cell is drawn.
                Call Draw_Square(X * 10, Y * 10)
            End If
        Next Y
    Next X
End Sub

Function Draw_Square(X As Integer, Y As Integer)

'Purpose: The purpose of this function is to draw a red 10 by 10 pixel square at the position the function is given through
'X and Y co-ordinates.

'Variable Dictionary

'Variable Name      Scope           Type        Purpose
'X                  Draw_Square     Integer     This variable holds the X co-ordinate of the square being drawn.
'Y                  Draw_Square     Integer     This variable holds the Y co-ordinate of the square being drawn.

    Line (X, Y)-(X + 10, Y + 10), RGB(255, 0, 0), BF
End Function

Private Sub cmdStop_Click()

'Purpose: The purpose of this button is to pause the program's calculations and visual display when clicked.

    Started = False
    cmdStop.Visible = False
    cmdStart.Visible = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Purpose: The purpose of this subroutine is to switch the state of a cell, from dead to alive or vice versa, when the
'cell is clicked on.

'Variable Dictionary

'Variable Name      Scope           Type        Purpose
'MX                 Form_MouseDown  Integer     The purpose of this variable is to hold the X position of the mouse.
'MY                 Form_MouseDown  Integer     The purpose of this variable is to hold the Y position of the mouse.

    Dim MX As Integer
    Dim MY As Integer
    
    MX = X / 10
    MY = Y / 10
    
    If MX < 60 And MY < 40 Then
        Alive(60 * MY + MX) = Not Alive(60 * MY + MX)   'Changes the cell's state to the opposite.
    End If
    
    Call Render
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Purpose: The purpose of this subroutine is to find the X and Y positions of the mouse, check if the position is inside
'the visual display, and draw a black square to indicate which cell's state will be changed if the mouse button is
'clicked.

'Variable Dictionary

'Variable Name      Scope           Type        Purpose
'MX                 Form_MouseMove  Integer     The purpose of this variable is to hold the X position of the mouse.
'MY                 Form_MouseMove  Integer     The purpose of this variable is to hold the Y position of the mouse.

    Dim MX As Integer
    Dim MY As Integer
    
    MX = X / 10         'Changes the mouse's position in pixels to the mouse's position in cell squares.
    MY = Y / 10
    
    If MX < 60 And MY < 40 Then         'If the mouse is in the visual display, a black square is drawn to represent it.
        Call Render
        Line (MX * 10, MY * 10)-((MX * 10) + 10, (MY * 10) + 10), RGB(0, 0, 0), BF
    End If
    
End Sub

Private Sub hsbTimerScroll_Change()

'Purpose: The purpose of this scroll bar is to allow the user to change the time interval between calculations, which
'slows down or speeds up the visual display.

    timLife.Interval = hsbTimerScroll.Value
End Sub

Private Sub timLife_Timer()

'Purpose: The purpose of this timer is to figure out which cells will be alive for the next calculation, and which will
'be dead.

'Variable Dictionary

'Variable Name      Scope       Type        Purpose
'X                  timLife     Integer     This variable is used as a counter in a For Loop, where it holds the X value,
'                                           or X position, of a cell, so that the cell, and its neighbours, can be found
'                                           by the program.
'Y                  timLife     Integer     This variable is used as a counter in a For Loop, where it holds the Y value,
'                                           or Y position, of a cell, so that the cell, and its neighbours, can be found
'                                           by the program.
'AliveCount         timLife     Integer     This variable counts how many of the cell's neighbours are alive, so that the
'                                           program can determine whether the cell will be alive or dead after the
'                                           calculation.


    Dim X As Integer
    Dim Y As Integer
    Dim AliveCount As Integer
    If Started = True Then          'If the Start button is clicked, and Started is true, the calculations run.
        For X = 0 To 59
            For Y = 0 To 39
                AliveCount = 0          'Resets AliveCount for current cell.
            
                If Is_Alive(X - 1, Y - 1) = True Then      'Uses IsAlive Function to check how many of the cell's 8
                    AliveCount = AliveCount + 1            'neighbours are alive.
                End If
                If Is_Alive(X, Y - 1) = True Then
                    AliveCount = AliveCount + 1
                End If
                If Is_Alive(X + 1, Y - 1) = True Then
                    AliveCount = AliveCount + 1
                End If
                If Is_Alive(X - 1, Y) = True Then
                    AliveCount = AliveCount + 1
                End If
                If Is_Alive(X + 1, Y) = True Then
                    AliveCount = AliveCount + 1
                End If
                If Is_Alive(X - 1, Y + 1) = True Then
                    AliveCount = AliveCount + 1
                End If
                If Is_Alive(X, Y + 1) = True Then
                    AliveCount = AliveCount + 1
                End If
                If Is_Alive(X + 1, Y + 1) = True Then
                    AliveCount = AliveCount + 1
                End If
            
                If (AliveCount = 2 Or AliveCount = 3) And Is_Alive(X, Y) Then 'Checks AliveCount with the game of life
                    NextAlive(60 * Y + X) = True                              'and figures out the cell's state.
                ElseIf AliveCount = 3 And Not Is_Alive(X, Y) Then
                    NextAlive(60 * Y + X) = True
                Else
                    NextAlive(60 * Y + X) = False
                End If
            
            Next Y
        Next X
        
        'Updates the Alive array using the NextAlive array.
        
        For X = 0 To 59
            For Y = 0 To 39
                If NextAlive(60 * Y + X) = True Then
                    Alive(60 * Y + X) = True
                ElseIf NextAlive(60 * Y + X) = False Then
                    Alive(60 * Y + X) = False
                End If
            Next Y
        Next X
    
        Call Render

    End If
    
End Sub

Function Is_Alive(X As Integer, Y As Integer) As Boolean

'Purpose: The purpose of this function is to check whether a given cell is alive or dead, and whether the X and Y
'co-ordinates given even correspond to a cell, so that the program doesn't crash at the borders of the visual display.

'Variable Dictionary

'Variable Name      Scope       Type        Purpose
'X                  Is_Alive    Integer     This variable holds the X value of the cell being checked.
'Y                  Is_Alive    Integer     This variable holds the Y value of the cell being checked.

    Is_Alive = False
    If X >= 0 And X < 60 And Y >= 0 And Y < 40 Then     'Makes sure the cell's position exists.
            If Alive(60 * Y + X) = True Then        'Checks if the cell is alive or dead
                Is_Alive = True
            End If
    End If
End Function
    
