Attribute VB_Name = "basFadePics"
Public Const FADE_T_TO_B = 0
Public Const FADE_B_TO_T = 1
Public Const FADE_L_TO_R = 2
Public Const FADE_R_TO_L = 3
Public Const FADE_RANDOM = 4
Public Const FADE_OUTWARD = 5

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub Fade(Pic As PictureBox, Style As Integer, Blocks As Integer)
   
    Dim width_section_size As Integer
    Dim height_section_size As Integer
    Dim i As Integer, j As Integer
    Dim save_color As Long
   
    'Saves the picbox's current forecolor
    save_color = Pic.ForeColor

    'Set Pics forecolor to its backcolor
    Pic.ForeColor = Pic.BackColor

    'Corrects the Blocks if needed
    If Blocks < 5 Then Blocks = 5
    If Blocks > 100 Then Blocks = 100

    'Sets the size of each width section
    width_section_size = Pic.ScaleWidth / Blocks

    'Sets the size of each height section
    height_section_size = Pic.ScaleHeight / Blocks


    Select Case Style
       '-------------------------------------------------------------------------------------
       Case 0  'Fading top to bottom
         
          For i = 0 To Blocks
             For j = 0 To Blocks
                Pic.Line ((j * width_section_size), (i * height_section_size))-((j + 1) * width_section_size, (i + 1) * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 1  'Fading bottom to top
         
          For i = Blocks To 0 Step -1
             For j = 0 To Blocks
                Pic.Line (((j - 1) * width_section_size), ((i - 1) * height_section_size))-(j * width_section_size, i * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 2  'Fading left to right
         
          For i = 0 To Blocks
             For j = 0 To Blocks
                Pic.Line ((i * width_section_size), (j * height_section_size))-((i + 1) * width_section_size, (j + 1) * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 3  'Fading right to left
         
          For i = Blocks To 0 Step -1
             For j = 0 To Blocks
                Pic.Line (((i - 1) * width_section_size), (j * height_section_size))-(i * width_section_size, (j + 1) * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 4  'Fading Random
       
          Dim bit_array() As Byte
          ReDim bit_array(Blocks, Blocks)
             
          Dim counter As Integer
       
          Do
             Do
                width_next_block = Int(Blocks * Rnd) 'Generate the random numbers
                height_next_block = Int(Blocks * Rnd) 'Generate the random numbers
                'MsgBox bit_array(width_next_block, height_next_block)
                If bit_array(width_next_block, height_next_block) = 0 Then
                  Exit Do
                End If
                counter = counter + 1
                If counter = Blocks * 10 Then Exit Do
             Loop
             
             If counter = Blocks * 10 Then Exit Do
             counter = 0
         
             'Update the bit_array
             bit_array(width_next_block, height_next_block) = 1
         
   
             
             Pic.Line ((width_next_block * width_section_size), (height_next_block * height_section_size))-((width_next_block + 1) * width_section_size, (height_next_block + 1) * height_section_size), , BF
         
             DoEvents
          Loop
         
          Pic.Line (0, 0)-(Pic.ScaleWidth, Pic.ScaleHeight), , BF
 
       '-------------------------------------------------------------------------------------
       Case 5 'Fading Outward
       
          For i = (Blocks / 2) To 0 Step -1
             Sleep (20)
             Pic.Line (i * width_section_size, i * height_section_size)-(((Blocks - i) + 1) * width_section_size, ((Blocks - i) + 1) * height_section_size), , BF
          Next
         
       '-------------------------------------------------------------------------------------
    End Select

    'Restores the picbox's original forecolor
    Pic.ForeColor = save_color
       
End Sub
