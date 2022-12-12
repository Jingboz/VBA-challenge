This is the homework of module 2 VBA

In this assignment, I splited the project into multiple functions for certain purposes, in order to be maintained easily

Sub init: this function is for setting up the column names for the next steps

Sub formatting: the conditional formatting function used for painting red and green in yearly change column

Sub sheetsLoop: this is the function for looping through each row in each sheet to pickup the year start and finish figures and manipulate for the dataset we are seeking.

Sub bonus: this function can realize the bonus requiremnets. Intead of using loops to go through all datasets from the previous steps, I utilized the inbuilt sorting function to find the max and min figures, which can save us heaps of time and computer memories. If I have more time on this, I'll try to implement arraylist or create a new class for the dataset and do sorting or ranking, which should provide a more robust program

Sub main: this is the main function to assemble all other functions and run them all at one time
