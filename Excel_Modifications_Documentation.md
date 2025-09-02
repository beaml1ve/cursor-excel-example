# Excel Modifications Documentation

## nola-ride-income.xlsx

**Date**: December 2024  
**File**: nola-ride-income.xlsx  
**Password**: beamL1ve

---

## 📋 Overview

This document describes all modifications made to the `nola-ride-income.xlsx` file to implement conditional fee calculations based on vehicle type (cab vs non-cab) and other enhancements.

---

## 🔧 Major Modifications

### 1. Credit Card Processing Fee Update

**Location**: `service`, `ecosystem`, `org` sheets  
**Change**: Updated credit card processing percentage from 2.9% to 3.51%  
**Affected Cells**:

- `service!B5`: 0.029 → 0.0351
- `ecosystem!B10`: 0.029 → 0.0351
- `org!B10`: 0.029 → 0.0351

### 2. Hail Ride Fee Adjustment

**Location**: `ecosystem` sheet  
**Change**: Updated hail ride fee from 1.5% to 2.51%  
**Affected Cell**: `ecosystem!B4`: 0.015 → 0.0251

### 3. Vehicle Type Conditional Logic Implementation

**Location**: `ecosystem` and `ecosystemFee` sheets  
**Purpose**: Different fee calculations for cab vs non-cab vehicles

#### 3.1 Ecosystem Sheet Modifications

**Original Categories** (rows 4-7):

- hail: 1.5%
- cad: 3.5%
- account: 2.0%
- app: 3.5%

**New Structure** (rows 4-11):

- hail: 1.5% (for non-cab vehicles)
- cad: 3.5% (for non-cab vehicles)
- account: 2.0% (for non-cab vehicles)
- app: 3.5% (for non-cab vehicles)
- hail_cab: 3.0% (doubled for cab vehicles)
- cad_cab: 7.0% (doubled for cab vehicles)
- account_cab: 4.0% (doubled for cab vehicles)
- app_cab: 7.0% (doubled for cab vehicles)

#### 3.2 EcosystemFee Sheet Restructure

**Purpose**: Implement conditional calculations based on `rideData!B51` (vehicleType)

**New Structure**:

- **Row 1**: Total calculation formula
- **Rows 4**: "Ride type fee components" header (merged A4:E4)
- **Rows 5-20**: 8 ride type categories (4 base × 2 vehicle types)
- **Row 21**: "Payment type fee components" header (merged A21:E21)
- **Rows 22-31**: 5 payment type categories (unchanged)

**Ride Type Categories**:

1. **Rows 5-6**: hail (non-cab) - `IF(AND(C5="hail",rideData!B$51<>"cab"),B6*C6+D6,0)`
2. **Rows 7-8**: hail_cab - `IF(AND(C7="hail",rideData!B$51="cab"),B8*C8+D8,0)`
3. **Rows 9-10**: cad (non-cab) - `IF(AND(C9="cad",rideData!B$51<>"cab"),B10*C10+D10,0)`
4. **Rows 11-12**: cad_cab - `IF(AND(C11="cad",rideData!B$51="cab"),B12*C12+D12,0)`
5. **Rows 13-14**: account (non-cab) - `IF(AND(C13="account",rideData!B$51<>"cab"),B14*C14+D14,0)`
6. **Rows 15-16**: account_cab - `IF(AND(C15="account",rideData!B$51="cab"),B16*C16+D16,0)`
7. **Rows 17-18**: app (non-cab) - `IF(AND(C17="app",rideData!B$51<>"cab"),B18*C18+D18,0)`
8. **Rows 19-20**: app_cab - `IF(AND(C19="app",rideData!B$51="cab"),B20*C20+D20,0)`

### 4. Table Optimization

**Location**: `ecosystemFee` sheet  
**Changes**:

- Removed 8 unnecessary padding rows between ride type categories
- Compacted structure for better organization
- Updated all formula references after row removal
- Merged header cells (A4:E4 and A21:E21)

### 5. Sheet Organization

**Change**: Moved `ecosystemFee` sheet to position 6 (after `serviceFee`)  
**New Sheet Order**:

1. rideData
2. service
3. ecosystem
4. org
5. serviceFee
6. ecosystemFee _(moved here)_
7. orgFee
8. userFee
9. result

### 6. Security Enhancement

**Change**: Protected all 9 sheets with password "beamL1ve"  
**Protected Sheets**: All sheets now require password for modifications

---

## 🎯 Key Logic Implementation

### Vehicle Type Detection

**Source**: `rideData!B51` (vehicleType field)  
**Logic**:

- If `vehicleType = "cab"` → Use doubled rates from `*_cab` categories
- If `vehicleType ≠ "cab"` → Use standard rates from base categories

### Conditional Formula Structure

```excel
=IF(AND(C[row]="[ride_type]",rideData!B$51="cab"),B[fee_row]*C[fee_row]+D[fee_row],0)
```

### Total Calculation

**Location**: `ecosystemFee!A1`  
**Formula**: `=ROUND(SUM(E6,E8,E10,E12,E14,E16,E18,E20,E23,E25,E27,E29,E31),2)`  
**Components**:

- E6,E8,E10,E12,E14,E16,E18,E20: Ride type calculations
- E23,E25,E27,E29,E31: Payment type calculations

---

## 📊 Data Flow

### Input Data

- **rideData!B8**: Ride type (hail, cad, account, app)
- **rideData!B51**: Vehicle type (cab, taxi, etc.)
- **rideData!B71**: Ride amount ($)

### Processing

1. **ecosystemFee** sheet checks ride type and vehicle type
2. Applies appropriate percentage and fixed fee from **ecosystem** sheet
3. Calculates: `(ride_amount × percentage) + fixed_fee`
4. Sums all applicable fees

### Output

- **ecosystemFee!A1**: Total ecosystem fee
- **result!B4**: References ecosystemFee calculation

---

## 🔍 Testing Scenarios

### Scenario 1: Cab Vehicle, Hail Ride, $100

- **Expected**: $3.00 (3% of $100)
- **Active Formula**: hail_cab calculation
- **Rate Used**: ecosystem!B8 (3.0%)

### Scenario 2: Non-Cab Vehicle, Hail Ride, $100

- **Expected**: $1.50 (1.5% of $100)
- **Active Formula**: hail non-cab calculation
- **Rate Used**: ecosystem!B4 (1.5%)

### Scenario 3: Cab Vehicle, CAD Ride, $100

- **Expected**: $7.00 (7% of $100)
- **Active Formula**: cad_cab calculation
- **Rate Used**: ecosystem!B9 (7.0%)

---

## ⚠️ Important Constraints

### Protected Areas

- **rideData sheet**: Format cannot be changed, only "default" row content can be modified
- **result sheet**: Format cannot be changed, formulas in B3-B7 can be modified
- **No row/column insertion or removal** allowed in rideData and result sheets

### Calculation Dependencies

- **Fee sheets** (serviceFee, ecosystemFee, orgFee, userFee) calculate output into A1 cells
- **result!B4:B7** reference these A1 calculations via formulas
- **ecosystem, service, org sheets** contain the fee conditions and percentages

---

## 🛠️ Technical Implementation Details

### Formula Pattern for Ride Types

```excel
Category Row: =ecosystem!$A$[row]
Fee Row A: fee
Fee Row B: =rideData!B$71 (ride amount)
Fee Row C: =ecosystem!$B$[row] (percentage)
Fee Row D: =ecosystem!$C$[row] (fixed fee)
Fee Row E: =IF(AND(C[cat_row]="[type]",rideData!B$51[condition]),B[fee_row]*C[fee_row]+D[fee_row],0)
```

### Ecosystem References

- **Non-cab categories**: ecosystem!$A$4:$A$7 (hail, cad, account, app)
- **Cab categories**: ecosystem!$A$8:$A$11 (hail_cab, cad_cab, account_cab, app_cab)

### Cell Formatting

- **Headers**: Merged across columns A-E
- **Category rows**: Light gray background (#EEEEEE)
- **Fee calculation columns**: Currency format for D and E columns

---

## 📝 Version History

1. **Initial State**: Basic fee calculation without vehicle type conditions
2. **Credit Card Update**: Changed processing fee to 3.51%
3. **Hail Fee Update**: Changed hail fee to 2.51%
4. **Conditional Logic**: Implemented vehicle type-based calculations
5. **Table Optimization**: Removed padding rows and merged headers
6. **Final Organization**: Moved sheets and applied protection

---

## 🔐 Security

- **Password**: beamL1ve
- **Protection Level**: All sheets protected
- **Access**: Requires password for any modifications
- **Backup**: Original file preserved as nola-ride-income-original.xlsx

---

_This documentation reflects the final state of the Excel file after all modifications have been completed and tested._
