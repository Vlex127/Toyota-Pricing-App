# Toyota Pricing App

A Visual Basic 6 desktop application for calculating Toyota vehicle pricing with customizable options and import duty calculations.

## Overview

This application provides a user-friendly interface for calculating the total cost of Toyota vehicles based on various factors including engine size, manufacturing year, optional facilities, and import duty percentages.

## Features

- **Vehicle Configuration**: Input vehicle specifications including make, engine size, and manufacturing year
- **Optional Facilities**: Select additional features like air conditioning and roof rack
- **Import Duty Calculation**: Apply custom import duty percentages
- **Real-time Pricing**: Instant calculation of total vehicle cost
- **Data Validation**: Input validation to prevent calculation errors
- **Team Information**: Built-in team member directory

## Application Structure

### Main Components

- **frmSplash.frm**: Splash screen that displays for 3 seconds on application startup
- **Toyota Pricing App.frm**: Main application form with pricing calculator
- **Toyota Pricing App.vbp**: Visual Basic project file
- **Toyota Pricing App.exe**: Compiled executable application

### Key Form Elements

#### Main Calculator Form (CSC_207_GROUP_3)
- **Input Fields**:
  - Make (Text box - pre-filled with "Toyota")
  - Engine Size (Text box - numeric input)
  - Year Manufactured (Text box - numeric input)
  - Import Duty % (Text box - numeric input)

- **Optional Facilities**:
  - Air Conditioning (Checkbox - ₦50,000)
  - Roof Rack (Checkbox - ₦150,000)

- **Control Buttons**:
  - Calculate: Performs price calculation
  - Clear: Resets all input fields
  - Info: Displays team member information

- **Display**:
  - Result label showing final price in Euro format

## Pricing Logic

### Base Pricing
- **Base Price**: ₦5,000,000

### Engine Size Adjustment
- If engine size > 2.0L: + ₦500,000

### Year Adjustment
- If manufacturing year < 2015: - ₦200,000

### Optional Facilities
- Air Conditioning: + ₦50,000
- Roof Rack: + ₦150,000

### Import Duty
- Applied as percentage of total cost after all other adjustments

## Calculation Formula

```
Total Cost = Base Price + Engine Adjustment + Year Adjustment + Facilities Cost
Final Price = Total Cost + (Total Cost × Import Duty % / 100)
```

## Installation and Usage

### Requirements
- Windows operating system
- Visual Basic 6 runtime files (if running from source)
- Approximately 500KB disk space

### Running the Application
1. Double-click `Toyota Pricing App.exe` to launch
2. Wait for splash screen (3 seconds)
3. Enter vehicle specifications in the main form
4. Select optional facilities if desired
5. Enter import duty percentage
6. Click "Calculate" to see the final price
7. Use "Clear" to reset all fields for new calculations

## Development Team

This application was developed by CSC 207 Group 3, consisting of 34 students:

**Page 1 of 2:**
- Eniola Uthman Oluwadarasimi (240591081)
- EZE C.O (240591082)
- Ezebube Chizitelum Odirachukwunma (240591083)
- Ezeh Johncharles (240591084)
- Olawale Anuoluwapo Ezekiel (240591085)
- Fajemisin Alexander Olumuyiwa (240591087)
- Fakoya Khalid (240591088)
- Familusi Ayomikun (240591089)
- Fasasi Mayowa (240591090)
- Fasoyiro Samuel Oreoluwa (240591091)
- Giwa Adesope Abdulmalik (240591092)
- Giwa Aliameen Opeyemi (240591093)
- Marvellous Godwin Osemudiamen (240591094)
- Hammed Abdulrahman Adesina (240591095)
- Hanidu Olawale Murtadha (240591096)
- Hassan Mubarak Atanda (240591097)
- Segun Ganiu Hassan (240591098)
- Idowu Fawaz Olawunmi (240591101)

**Page 2 of 2:**
- Idowu Matthew (240591102)
- Ikechukwu Queen (240591103)
- Inedu Esther Ogaicha (240591104)
- Iniotuh Greatest (240591105)
- Iwakun Oluwasegun Omotojumi (240591106)
- Iwuno Vincent ChukwuEbuka (240591107)
- Iyaniwura Olabamiji George (240591108)
- Jawando Fuad Olamide (240591109)
- Jeremiah David Preye (240591110)
- Joseph Elizabeth Nifemi (240591111)
- Kalejaiye Halimah Temilade (240591112)
- Kasali Damilola Emmanuel (240591113)
- Kehinde Oyindamola Ayomide (240591115)
- Kolawole Abubakar Olaoluwa (240591116)
- Kwegan Sean Oluwatomilade (240591117)
- Lamina Rihanat Opemipo (240591118)
- Lawal Sahal Adeshayo (240591119)
- Ibrahim Oluwafemi Lawal (240591120)

## Technical Specifications

- **Development Environment**: Visual Basic 6
- **Application Type**: Windows Forms Application
- **Project Version**: 1.0.0
- **Company**: vlex
- **Compilation**: Native code with no debugging information

## File Structure

```
Toyota Pricing App/
├── README.md                     # This documentation file
├── Toyota Pricing App.exe        # Compiled executable
├── Toyota Pricing App.frm        # Main form source code
├── Toyota Pricing App.frx        # Form resources (images, icons)
├── Toyota Pricing App.vbp        # Project file
├── Toyota Pricing App.vbw        # Workspace settings
├── frmSplash.frm                 # Splash screen form
└── frmSplash.frx                # Splash screen resources
```

## Error Handling

The application includes comprehensive input validation:
- Numeric validation for engine size, year, and duty percentage
- Prevents calculation crashes from invalid input
- User-friendly error messages for invalid data

## Future Enhancements

Potential areas for improvement:
- Support for multiple vehicle makes beyond Toyota
- Database integration for vehicle specifications
- Advanced depreciation calculations
- Currency conversion features
- Export functionality for calculation results

## License

This project was developed as part of CSC 207 coursework. Please contact the development team for usage permissions.