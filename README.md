# AutoFilterManagerForExcel
Classes and implementation to record status of autofilters in tables or a range in a worksheet, remove the filtering, and then restore the filters

<b>Background and description of the solution developed</b>

Bill Benson on the Excel-L list group posed the initial problem with the subject "High Quality Code for Capturing and re-Applying Autofilters". What I came up with is available in this Repository. I cannot vouch for the "High Quality" part of the solution I came up with, but it does seem to answer the mail; only Bill can attest to that.

Since I use Rubberduck (latest version here: https://github.com/rubberduck-vba/Rubberduck/releases) it is easier to navigate and understand the project when viewd through the RD lens. I suggest you get it and find out how truly useful it is. Help is available in the the Rubberducking chat room: https://chat.stackexchange.com/rooms/14929/vba-rubberducking and the Rubberduck web site: https://rubberduckvba.com/

There are 3 sheets in the workbook: 1) RangeSolution; 2) TableSolution; 3) SheetSolution.
The 3 sheets mirror the development approach. I got the RangeSolution working first, which fleshed out and solved the fundamental problem. Expanding to do a table was a slight modification to deal with tables rather than ranges; it was pretty straightforward. Pulled it all together to handle multiple tables along with a range on a single sheet.

On each sheet there is a "Run Me" button to show the solution in action. Clicking the button stores the filter settings, removes the filtering, and pops up a message box to allow viewing of the unfiltered solution being explored. Dismissing the message box results in the filtering being restored.

In each solution the same pattern of filtering is used in the filtered starting points: Leftmost filter is a selection from the drop down selector box; the middle filter is a text filter using two strings with an "Or" operator; the rightmost filter is not applied.

There are 3 objects in the project: 1) AutoFilterManagerModel; 2) AutoFilterModel; 3) FilterModel

Entry points are provided for each of the 3 solutions illustrating how to set up and implement the models.

Thanks to Mathieu Guindon for sharing his techniques for using models and other ideas. I did not fully flesh this out with Interfaces or factory methods, but I can see where that could make this more robust. Nor is there any unit testing.

This could conveivably be transformed into an Add-In with a nice user interface to select workbooks, worksheets, filtered ranges and tables to be processed. Creating a full blown solution was not really the point. However, if anyone is interested in advancing this project please use this repository to do that work so that more people can benefit from those efforts.

Regards,

David G. Miley
