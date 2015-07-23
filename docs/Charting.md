#Charting
The bUTL add-in really shines when it comes to its charting features.  These work by enhancing the existing charting features, but it automates a large part of the usually tedious process.

Interface for charting features.
##Create grid of charts
This button will open a form that allows you to configure a grid of charts.  It applies to all the charts on the current sheet.  The charts are arranged in the order that they were added to the sheet.  If you want to reorder the charts, you can cut and paste the chart on to put it at the end of the list.

Images show before and after grid creation.

![grid before](.\images\charting\grid before.png)

![grid after](.\images\charting\grid after.png)

The options in this menu do the following things:

 - Columns refers to the number of columns of charts
 - Height and width refer to each individual chart.  They will all be made the same size.  These numbers are in pixels
 - The v and h offsets are the vertical and horizontal offsets from the edges of the sheet.  If these are 0 and 0, the charts will be in the corner of the sheet.
 - Down first is a check box that allows the charts to run down the page first before they go across,
 - Zoom in after is an option that will change the zoom so that all of the charts fit horizontally.
##Format all charts
This button will apply a generic formatting to all of the charts in the current sheet.  It is made for scatter plot type charts.  The formatting options currently include:

 - Sets the marker style to a circle (instead of automatic) and the size to 3
 - If there is a line connecting dots, it will make this 1.5 wide
 - Sets the color to automatic for the marker and removes the second “border” color from the marker.
 - Adds a legend to the chart and puts it at the bottom of the chart
 - Changes the colors for the grid lines to be a light grey
 - If the chart has a title, it sets the format to bold and the font size to 12.

##Create set of time series charts
This option will allow you to create a set of charts that use a common x-axis and different y-axes.  It also allows you to select the range of cells to use for the series names/chart titles.  This works well to create a block of charts that plot a variable vs. time. This can also be used to make a series of charts that have a common x-axis.  This works well for creating a block of charts that can then be used with other features.  After hitting the button, there are 3 menus to make the range selections.
Select the date range (or whatever you want on the x-axis)
Select the data (needs to be the same height as the dates)
Choose the titles (needs to be the same width as the data)

![time series](.\images\charting\time series.png)

##Create set of XY matrix charts
This option will produce a matrix of XY charts that show the relationship/correlation between variables.  It will create all of the charts in one shot which makes it preferable to the manual option if there are a large number of variables.
Select the data with the titles in the top row
This will create a box of scatterplots.

![xy matrix](.\images\charting\xy matrix.png)

##Flip XY associations
This works on the selected charts.  It will flip the X and Y series.  If a chart has multiple series on it, it will flip all of them.  This works well for XY scatterplots in order to flip the relationship.  It will also flip the X and Y axis labels if any exist.

![flip xy](.\images\charting\flip xy.png)

##Series options
This contains some features dedicated to the series on a chart.
###Merge series
This works on selected charts.  It will merge the series from one chart and put them into another chart.  This will erase the formatting for one set of series.  This can then be overcome by using the format all button above.

![series merge before](.\images\charting\series merge before.png)

![series merge after](.\images\charting\series merge after.png)

###Split series
This works on selected charts.  It will take all of the series and put each one in its own chart.  This can be combined with the grid to then display them all correctly.  This is best used to split all the series out so that they can be merged back together differently.

![series split before](.\images\charting\series split before.png)

![series split after](.\images\charting\series split after.png)

##Find Data menu
This menu shows the different options to find the data within a series.
###Find X series range
The find X option is used to select the data that is being used for the X axis on a chart.  It applies to the currently selected series.
###Find Y series range
This is the same as the find X option except it works on the y axis.
##Fit Data menu/button
This menu is used to change the behavior of the axes ranges for all of the axes on a chart.  It can be used to fit the axes to the data or to revert back to the default AUTO range.  These work on all of the currently selected charts.  Hitting this button (instead of opening the menu) will fit both the x and y axis data.
###Fit X / Fit Y
These will fit the axis range to the min and maximum of the data set.  It will include hidden cells even if they are not in the chart.

![fit xy before](.\images\charting\fit x y before.png)

![fit x after](.\images\charting\fit x y after x.png)

![fit y after](.\images\charting\fit x y after both.png)

###Auto X / Auto Y
These will revert the axes back to the default AUTO setting.
##Extend series
This will extend the series for the currently selected charts.  This works if you have added data below the current range of the chart.  Hitting this button will cause the range of series to be set to the “full” range of the current data.  This is the same as doing “edit data” and then selecting the full block of data again.

![extend series](.\images\charting\extend series.png)

##Add Trendlines
This will add a unique trend line to each series on the chart.  By default, it adds a linear trend to the chart.  In order to aid identification, it will also color the series using the bUTL color scheme.  The trend line labels will then be colored the same as the series.  There is a technicality with the auto colors that does not allow them to be easily identified in code; therefore I have to force the color in order to determine it.

![add trendlines](.\images\charting\add trendlines.png)

##Apply Colors
This will apply the bUTL color scheme to the currently selected charts.  This scheme includes 10 colors that are more distinct (and visual appealing) than the default Office 2010 colors.

![apply colors](.\images\charting\apply colors.png)

##Add Titles menu/button
This menu contains options for adding titles to the currently selected charts.  These are usually faster than clicking several times on the normal Chart ribbon menus.  Hitting this button (instead of opening the menu) will add dummy titles to the chart title, y-axis, and x-axis.  The idea is that it is easier to change this text or delete an extra axis then to click through all the menus to add these normally.

![add titles](.\images\charting\add titles.png)

###Add title by series
This will set the y-axis title to be the same as series name for each chart that is selected.  For a large number of charts, this makes it easy to use the series name and then update the axes to say what you want.

![add titles by series before](.\images\charting\add title by series before.png)

![add titles by series after](.\images\charting\add title by series after.png)
