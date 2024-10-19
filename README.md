# AutoSafetyTools
 Various automation tools for risk & safety assessments

## WindRose
### Creates and displays wind rose plots from wind data.
- Import .dlp file of weather data.
- Generate wind rose image and XML file for MERIT input.
- Export image and XML

## ShepherdAnalyser
### Analyzes exceedance and risk values from Shell Shepherd.
- Import Shell Shepherd risk results XML.
- Uses basic image recognition to analyse exceedance values at specified risk values.
- Export Excel file of exceedance values of each building.

## SCEAnalyser
### GUI for analyzing Safety Critical Elements from Excel files from BowTieXP.
- Import BowTieXP "Barrier Register" Report.
- Generate Excel worksheet of summary of applicable SCEs.
- Export Excel worksheet.

## GridGenerator
### Create grid images.
- Set image & grid dimensions, background color, grid line type, weight, and color.
- Generate grid image.
- Export grid image.

## ProbitCalculator
### Calculates lethality values from Probit constants.
- Models lethality from toxic exposure as function of time and concentration based on Probit constants.
- Generate 3D plots of lethality, time and concentration.
- No export functionality.

## PartsCountXML
### Generates XML for parts count data for MERIT input.
- Input Excel sheet of parts count data.
- Export XML of parts count for MERIT input.
