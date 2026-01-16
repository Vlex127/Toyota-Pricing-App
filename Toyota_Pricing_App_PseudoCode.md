# Toyota Pricing Application - Pseudo Code

## Application Overview
A Visual Basic application that calculates Toyota vehicle pricing based on engine size, manufacturing year, import duty, and optional features.

## Main Components

### 1. Splash Screen (frmSplash)
```
START Splash Screen:
    Display application title "Toyota Pricing App"
    Display subtitle "Professional Vehicle Pricing Solution"
    Display version "Version 1.0"
    Display logo/image
    Start timer (3 seconds)
    
    WHEN timer expires:
        Close splash screen
        Open main calculator form
END
```

### 2. Main Calculator Form (CSC_207_GROUP_3)
```
START Main Application:
    Initialize form with:
        - Input fields: Engine Size, Year, Import Duty %
        - Checkboxes: Air Conditioning, Open Roof
        - Buttons: Calculate Cost, Clear, Info
        - Result display label
        
    SET default values:
        - Make field = "Toyota"
        - Result = "N 0.00"
END
```

### 3. Calculate Cost Algorithm
```
FUNCTION Calculate_Total_Cost():
    // Input Validation
    IF engine_size is not numeric OR empty:
        SHOW error "Please enter valid Engine Size"
        FOCUS on engine field
        EXIT function
        
    IF year is not numeric OR empty:
        SHOW error "Please enter valid Year"
        FOCUS on year field
        EXIT function
        
    IF duty_percent is not numeric OR empty:
        SHOW error "Please enter valid Duty %"
        FOCUS on duty field
        EXIT function
    
    // Get input values
    engine = CONVERT_TO_NUMBER(engine_field)
    year = CONVERT_TO_INTEGER(year_field)
    duty = CONVERT_TO_NUMBER(duty_field)
    
    // Calculate base price by engine size
    IF engine <= 2000:
        base_price = 1,500,000  // Small engine
    ELSE IF engine <= 3000:
        base_price = 2,500,000  // Medium engine
    ELSE:
        base_price = 4,000,000  // Large engine
    
    // Adjust base price by year
    IF year >= 2020:
        base_price = base_price + 500,000  // Newer car premium
    ELSE IF year < 2015:
        base_price = base_price - 300,000  // Older car discount
    
    // Calculate optional features cost
    facilities_cost = 0
    IF air_conditioning_checked:
        facilities_cost = facilities_cost + 75,000
    
    IF open_roof_checked:
        facilities_cost = facilities_cost + 50,000
    
    // Calculate import duty
    import_duty = base_price * (duty / 100)
    
    // Calculate total cost
    total_cost = base_price + import_duty + facilities_cost
    
    // Display formatted result
    DISPLAY "N " + FORMAT(total_cost, "#,##0.00")
    
END FUNCTION
```

### 4. Clear Function
```
FUNCTION Clear_Form():
    SET engine_field = ""
    SET year_field = ""
    SET duty_field = ""
    UNCHECK air_conditioning
    UNCHECK open_roof
    SET result_label = "N 0.00"
    FOCUS on engine_field
END FUNCTION
```

### 5. Info/About Function
```
FUNCTION Show_Information():
    // Display 3-page information dialog
    
    PAGE 1 - Usage Instructions:
        - How to use the application
        - Pricing logic explanation
        - First 10 group members
    
    PAGE 2 - Group Members:
        - Next 14 group members
    
    PAGE 3 - Group Members:
        - Final 12 group members
        - Thank you message
END FUNCTION
```

## Application Flow
1. **Startup**: Show splash screen for 3 seconds
2. **Main Interface**: Display calculator form
3. **User Input**: Enter vehicle specifications
4. **Calculation**: Click "Calculate Cost" button
5. **Validation**: Check all inputs are valid numbers
6. **Processing**: Apply pricing algorithm
7. **Display**: Show formatted total cost
8. **Options**: Clear form or view information

## Data Structures
- **Variables**: BasePrice, EngineSize, YearMade, DutyPercent, FacilitiesCost, TotalCost
- **Controls**: Text boxes, check boxes, command buttons, labels
- **Constants**: Feature costs (AC: N75,000, Roof: N50,000)

## Error Handling
- Input validation for all numeric fields
- User-friendly error messages
- Focus management for invalid fields
- Graceful exit on validation errors

## Pricing Logic Details

### Base Price by Engine Size:
- Small engines (≤2000cc): N1,500,000
- Medium engines (≤3000cc): N2,500,000
- Large engines (>3000cc): N4,000,000

### Year Adjustments:
- Newer cars (≥2020): +N500,000
- Older cars (<2015): -N300,000

### Optional Features:
- Air Conditioning: +N75,000
- Open Roof: +N50,000

### Import Duty:
- Calculated as percentage of base price
- User-defined percentage rate

### Final Calculation:
Total Cost = Base Price + Year Adjustment + Import Duty + Optional Features
