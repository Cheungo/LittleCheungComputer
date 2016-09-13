# LittleCheungComputer
My own spin of the Little Man Computer (http://peterhigginson.co.uk/LMC/) CPU simulation program - with a more simplistic and less "what is happening right now" user interface.

Little Cheung Computer was created for my OCR A Level Computing Coursework (F454) using Visual Basic. 

General Usage
=============
Open program by double clicking "Little Cheung Computer.exe"

Assembly code (Refer to Instruction Set in "Help" drop down menu) can be entered in the left text box.

To execute the program, press the "Run" button.

The "Program Counter" keeps count of the number of steps of the program cycle. It is used in this program for preventing programs with too many cycles - e.g. infinite loops.

The "Accumulator" stores temporary results of arithmetic calculations.

"Input" and "Output" displays the input and output.


Example program:
INP (Input - enter an input in the input box. Input is stored in the accumulator; just like in a CPU)
OUT (Output - displays the output in the output box. Output is retrieved from the accumulator)
HLT (Halt - stops the program)


The text boxes on the right are "Registers". They are used to hold temporary data to be used when processing instructions - including storing the actual instruction. 

For example, INP, on line 0, has machine code 901. Since the instruction is on line 0, so 901 is stored in Register 0.





