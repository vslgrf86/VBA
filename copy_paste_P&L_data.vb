
Sub SG_GROUP_copy_paste_jan_dez_values_from_monster()

'clear all existing values jan-dez
Range("C6:N41").Select
    Selection.ClearContents
    
    
'Jan copy values from monster P&L


'Sales SaaS Licenses
        Range("C6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C48/1000"
        
 'Sales SaaS Service
        Range("C7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C48/1000"
        
    'Sales Media Posting
        Range("C8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C48/1000"
        
        'Sales Commission
        Range("C9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C48/1000"
        
        'Total Sales
        Range("C10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C48/1000"
        
        'Media Posting
        Range("C11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C48/1000"
        
        'Others (Textkernel)
        Range("C12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C48/1000"
        
        'COGS
        Range("C13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C48/1000"
        
        'Other Income/(Expense)
        Range("C14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C48/1000"
        
        'Gross Profit
        Range("C15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C48/1000"
        
         'Employee Benefits
        Range("C17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C48/1000"
        
        'External services/Freelancer
        Range("C18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C48/1000"
        
        'Legal and Consulting Costs
        Range("C19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C48/1000"
      
        'Audit Costs
        Range("C20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C48/1000"
        
     'License Costs
        Range("C21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C48/1000"
        
         'Marketing Expenses
        Range("C22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C48/1000"
        
          'Travel Expenses
        Range("C23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C48/1000"
        
          'Car Expenses
        Range("C24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C48/1000"
        
          'Office Costs
        Range("C25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C48/1000"
        
        'Hosting Costs
        Range("C26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C48/1000"
        
        'Admin/Other Expenses
        Range("C27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C48/1000"
        
        'Total Operating Expenses
        Range("C28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C48/1000"
        
        'EBITDA
        Range("C29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C48/1000"
        
        'Depreciation
        Range("C30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C48/1000"
        
         'Capitalised Costs
        Range("C31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C48/1000"
        
        'Depreciation Capitalised Costs
        Range("C32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C48/1000"
        
        'M&A Expenses
        Range("C33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C48/1000"
        
         'EBIT
        Range("C34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C48/1000"
        
         'Interest Income
        Range("C35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C48/1000"
        
        'Interest Expense
        Range("C36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C48/1000"
        
        'Extraordinary Income
        Range("C37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C48/1000"
        
         'Extraordinary Expenses
        Range("C38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C48/1000"
        
        'Earnings Before Tax
        Range("C39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C48/1000"
        
        'Tax
        Range("C40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C48/1000"
        
         'Earnings After Tax
        Range("C41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C48/1000"
    
    

'Feb copy values from monster P&L


'Sales SaaS Licenses
        Range("D6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C49/1000"
        
 'Sales SaaS Service
        Range("D7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C49/1000"
        
    'Sales Media Posting
        Range("D8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C49/1000"
        
        'Sales Commission
        Range("D9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C49/1000"
        
        'Total Sales
        Range("D10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C49/1000"
        
        'Media Posting
        Range("D11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C49/1000"
        
        'Others (Textkernel)
        Range("D12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C49/1000"
        
        'COGS
        Range("D13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C49/1000"
        
        'Other Income/(Expense)
        Range("D14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C49/1000"
        
        'Gross Profit
        Range("D15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C49/1000"
        
         'Employee Benefits
        Range("D17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C49/1000"
        
        'External services/Freelancer
        Range("D18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C49/1000"
        
        'Legal and Consulting Costs
        Range("D19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C49/1000"
      
        'Audit Costs
        Range("D20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C49/1000"
        
     'License Costs
        Range("D21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C49/1000"
        
         'Marketing Expenses
        Range("D22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C49/1000"
        
          'Travel Expenses
        Range("D23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C49/1000"
        
          'Car Expenses
        Range("D24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C49/1000"
        
          'Office Costs
        Range("D25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C49/1000"
        
        'Hosting Costs
        Range("D26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C49/1000"
        
        'Admin/Other Expenses
        Range("D27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C49/1000"
        
        'Total Operating Expenses
        Range("D28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C49/1000"
        
        'EBITDA
        Range("D29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C49/1000"
        
         'Depreciation
        Range("D30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C49/1000"
        
         'Capitalised Costs
        Range("D31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C49/1000"
        
        'Depreciation Capitalised Costs
        Range("D32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C49/1000"
        
        'M&A Expenses
        Range("D33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C49/1000"
        
         'EBIT
        Range("D34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C49/1000"
        
         'Interest Income
        Range("D35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C49/1000"
        
        'Interest Expense
        Range("D36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C49/1000"
        
        'Extraordinary Income
        Range("D37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C49/1000"
        
         'Extraordinary Expenses
        Range("D38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C49/1000"
        
        'Earnings Before Tax
        Range("D39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C49/1000"
        
        'Tax
        Range("D40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C49/1000"
        
         'Earnings After Tax
        Range("D41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C49/1000"
        
        
        
'Mrz copy values from monster P&L


'Sales SaaS Licenses
        Range("E6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C50/1000"
        
 'Sales SaaS Service
        Range("E7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C50/1000"
        
    'Sales Media Posting
        Range("E8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C50/1000"
        
        'Sales Commission
        Range("E9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C50/1000"
        
        'Total Sales
        Range("E10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C50/1000"
        
        'Media Posting
        Range("E11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C50/1000"
        
        'Others (Textkernel)
        Range("E12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C50/1000"
        
        'COGS
        Range("E13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C50/1000"
        
        'Other Income/(Expense)
        Range("E14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C50/1000"
        
        'Gross Profit
        Range("E15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C50/1000"
        
         'Employee Benefits
        Range("E17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C50/1000"
        
        'External services/Freelancer
        Range("E18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C50/1000"
        
        'Legal and Consulting Costs
        Range("E19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C50/1000"
      
        'Audit Costs
        Range("E20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C50/1000"
        
     'License Costs
        Range("E21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C50/1000"
        
         'Marketing Expenses
        Range("E22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C50/1000"
        
          'Travel Expenses
        Range("E23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C50/1000"
        
          'Car Expenses
        Range("E24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C50/1000"
        
          'Office Costs
        Range("E25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C50/1000"
        
        'Hosting Costs
        Range("E26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C50/1000"
        
        'Admin/Other Expenses
        Range("E27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C50/1000"
        
        'Total Operating Expenses
        Range("E28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C50/1000"
        
        'EBITDA
        Range("E29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C50/1000"
        
         'Depreciation
        Range("E30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C50/1000"
        
         'Capitalised Costs
        Range("E31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C50/1000"
        
        'Depreciation Capitalised Costs
        Range("E32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C50/1000"
        
        'M&A Expenses
        Range("E33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C50/1000"
        
         'EBIT
        Range("E34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C50/1000"
        
         'Interest Income
        Range("E35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C50/1000"
        
        'Interest Expense
        Range("E36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C50/1000"
        
        'Extraordinary Income
        Range("E37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C50/1000"
        
         'Extraordinary Expenses
        Range("E38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C50/1000"
        
        'Earnings Before Tax
        Range("E39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C50/1000"
        
        'Tax
        Range("E40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C50/1000"
        
         'Earnings After Tax
        Range("E41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C50/1000"
        
        
        
        
        'Apr copy values from monster P&L


'Sales SaaS Licenses
        Range("F6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C51/1000"
        
 'Sales SaaS Service
        Range("F7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C51/1000"
        
    'Sales Media Posting
        Range("F8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C51/1000"
        
        'Sales Commission
        Range("F9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C51/1000"
        
        'Total Sales
        Range("F10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C51/1000"
        
        'Media Posting
        Range("F11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C51/1000"
        
        'Others (Textkernel)
        Range("F12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C51/1000"
        
        'COGS
        Range("F13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C51/1000"
        
        'Other Income/(Expense)
        Range("F14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C51/1000"
        
        'Gross Profit
        Range("F15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C51/1000"
        
         'Employee Benefits
        Range("F17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C51/1000"
        
        'External services/Freelancer
        Range("F18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C51/1000"
        
        'Legal and Consulting Costs
        Range("F19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C51/1000"
      
        'Audit Costs
        Range("F20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C51/1000"
        
     'License Costs
        Range("F21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C51/1000"
        
         'Marketing Expenses
        Range("F22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C51/1000"
        
          'Travel Expenses
        Range("F23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C51/1000"
        
          'Car Expenses
        Range("F24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C51/1000"
        
          'Office Costs
        Range("F25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C51/1000"
        
        'Hosting Costs
        Range("F26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C51/1000"
        
        'Admin/Other Expenses
        Range("F27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C51/1000"
        
        'Total Operating Expenses
        Range("F28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C51/1000"
        
        'EBITDA
        Range("F29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C51/1000"
        
         'Depreciation
        Range("F30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C51/1000"
        
         'Capitalised Costs
        Range("F31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C51/1000"
        
        'Depreciation Capitalised Costs
        Range("F32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C51/1000"
        
        'M&A Expenses
        Range("F33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C51/1000"
        
         'EBIT
        Range("F34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C51/1000"
        
         'Interest Income
        Range("F35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C51/1000"
        
        'Interest Expense
        Range("F36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C51/1000"
        
        'Extraordinary Income
        Range("F37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C51/1000"
        
         'Extraordinary Expenses
        Range("F38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C51/1000"
        
        'Earnings Before Tax
        Range("F39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C51/1000"
        
        'Tax
        Range("F40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C51/1000"
        
         'Earnings After Tax
        Range("F41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C51/1000"
        
        
        
        'Mai copy values from monster P&L


'Sales SaaS Licenses
        Range("G6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C52/1000"
        
 'Sales SaaS Service
        Range("G7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C52/1000"
        
    'Sales Media Posting
        Range("G8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C52/1000"
        
        'Sales Commission
        Range("G9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C52/1000"
        
        'Total Sales
        Range("G10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C52/1000"
        
        'Media Posting
        Range("G11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C52/1000"
        
        'Others (Textkernel)
        Range("G12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C52/1000"
        
        'COGS
        Range("G13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C52/1000"
        
        'Other Income/(Expense)
        Range("G14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C52/1000"
        
        'Gross Profit
        Range("G15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C52/1000"
        
         'Employee Benefits
        Range("G17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C52/1000"
        
        'External services/Freelancer
        Range("G18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C52/1000"
        
        'Legal and Consulting Costs
        Range("G19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C52/1000"
      
        'Audit Costs
        Range("G20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C52/1000"
        
     'License Costs
        Range("G21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C52/1000"
        
         'Marketing Expenses
        Range("G22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C52/1000"
        
          'Travel Expenses
        Range("G23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C52/1000"
        
          'Car Expenses
        Range("G24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C52/1000"
        
          'Office Costs
        Range("G25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C52/1000"
        
        'Hosting Costs
        Range("G26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C52/1000"
        
        'Admin/Other Expenses
        Range("G27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C52/1000"
        
        'Total Operating Expenses
        Range("G28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C52/1000"
        
        'EBITDA
        Range("G29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C52/1000"
        
         'Depreciation
        Range("G30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C52/1000"
        
         'Capitalised Costs
        Range("G31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C52/1000"
        
        'Depreciation Capitalised Costs
        Range("G32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C52/1000"
        
        'M&A Expenses
        Range("G33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C52/1000"
        
         'EBIT
        Range("G34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C52/1000"
        
         'Interest Income
        Range("G35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C52/1000"
        
        'Interest Expense
        Range("G36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C52/1000"
        
        'Extraordinary Income
        Range("G37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C52/1000"
        
         'Extraordinary Expenses
        Range("G38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C52/1000"
        
        'Earnings Before Tax
        Range("G39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C52/1000"
        
        'Tax
        Range("G40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C52/1000"
        
         'Earnings After Tax
        Range("G41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C52/1000"
        
        
        
        
        'Jun copy values from monster P&L


'Sales SaaS Licenses
        Range("H6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C53/1000"
        
 'Sales SaaS Service
        Range("H7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C53/1000"
        
    'Sales Media Posting
        Range("H8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C53/1000"
        
        'Sales Commission
        Range("H9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C53/1000"
        
        'Total Sales
        Range("H10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C53/1000"
        
        'Media Posting
        Range("H11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C53/1000"
        
        'Others (Textkernel)
        Range("H12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C53/1000"
        
        'COGS
        Range("H13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C53/1000"
        
        'Other Income/(Expense)
        Range("H14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C53/1000"
        
        'Gross Profit
        Range("H15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C53/1000"
        
         'Employee Benefits
        Range("H17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C53/1000"
        
        'External services/Freelancer
        Range("H18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C53/1000"
        
        'Legal and Consulting Costs
        Range("H19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C53/1000"
      
        'Audit Costs
        Range("H20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C53/1000"
        
     'License Costs
        Range("H21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C53/1000"
        
         'Marketing Expenses
        Range("H22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C53/1000"
        
          'Travel Expenses
        Range("H23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C53/1000"
        
          'Car Expenses
        Range("H24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C53/1000"
        
          'Office Costs
        Range("H25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C53/1000"
        
        'Hosting Costs
        Range("H26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C53/1000"
        
        'Admin/Other Expenses
        Range("H27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C53/1000"
        
        'Total Operating Expenses
        Range("H28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C53/1000"
        
        'EBITDA
        Range("H29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C53/1000"
        
         'Depreciation
        Range("H30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C53/1000"
        
         'Capitalised Costs
        Range("H31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C53/1000"
        
        'Depreciation Capitalised Costs
        Range("H32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C53/1000"
        
        'M&A Expenses
        Range("H33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C53/1000"
        
         'EBIT
        Range("H34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C53/1000"
        
         'Interest Income
        Range("H35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C53/1000"
        
        'Interest Expense
        Range("H36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C53/1000"
        
        'Extraordinary Income
        Range("H37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C53/1000"
        
         'Extraordinary Expenses
        Range("H38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C53/1000"
        
        'Earnings Before Tax
        Range("H39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C53/1000"
        
        'Tax
        Range("H40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C53/1000"
        
         'Earnings After Tax
        Range("H41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C53/1000"
        
        
        'Jul copy values from monster P&L


'Sales SaaS Licenses
        Range("I6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C54/1000"
        
 'Sales SaaS Service
        Range("I7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C54/1000"
        
    'Sales Media Posting
        Range("I8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C54/1000"
        
        'Sales Commission
        Range("I9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C54/1000"
        
        'Total Sales
        Range("I10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C54/1000"
        
        'Media Posting
        Range("I11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C54/1000"
        
        'Others (Textkernel)
        Range("I12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C54/1000"
        
        'COGS
        Range("I13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C54/1000"
        
        'Other Income/(Expense)
        Range("I14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C54/1000"
        
        'Gross Profit
        Range("I15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C54/1000"
        
         'Employee Benefits
        Range("I17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C54/1000"
        
        'External services/Freelancer
        Range("I18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C54/1000"
        
        'Legal and Consulting Costs
        Range("I19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C54/1000"
      
        'Audit Costs
        Range("I20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C54/1000"
        
     'License Costs
        Range("I21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C54/1000"
        
         'Marketing Expenses
        Range("I22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C54/1000"
        
          'Travel Expenses
        Range("I23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C54/1000"
        
          'Car Expenses
        Range("I24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C54/1000"
        
          'Office Costs
        Range("I25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C54/1000"
        
        'Hosting Costs
        Range("I26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C54/1000"
        
        'Admin/Other Expenses
        Range("I27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C54/1000"
        
        'Total Operating Expenses
        Range("I28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C54/1000"
        
        'EBITDA
        Range("I29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C54/1000"
        
         'Depreciation
        Range("I30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C54/1000"
        
         'Capitalised Costs
        Range("I31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C54/1000"
        
        'Depreciation Capitalised Costs
        Range("I32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C54/1000"
        
        'M&A Expenses
        Range("I33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C54/1000"
        
         'EBIT
        Range("I34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C54/1000"
        
         'Interest Income
        Range("I35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C54/1000"
        
        'Interest Expense
        Range("I36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C54/1000"
        
        'Extraordinary Income
        Range("I37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C54/1000"
        
         'Extraordinary Expenses
        Range("I38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C54/1000"
        
        'Earnings Before Tax
        Range("I39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C54/1000"
        
        'Tax
        Range("I40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C54/1000"
        
         'Earnings After Tax
        Range("I41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C54/1000"
        
        
        
        
        'Aug copy values from monster P&L


'Sales SaaS Licenses
        Range("J6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C55/1000"
        
 'Sales SaaS Service
        Range("J7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C55/1000"
        
    'Sales Media Posting
        Range("J8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C55/1000"
        
        'Sales Commission
        Range("J9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C55/1000"
        
        'Total Sales
        Range("J10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C55/1000"
        
        'Media Posting
        Range("J11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C55/1000"
        
        'Others (Textkernel)
        Range("J12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C55/1000"
        
        'COGS
        Range("J13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C55/1000"
        
        'Other Income/(Expense)
        Range("J14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C55/1000"
        
        'Gross Profit
        Range("J15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C55/1000"
        
         'Employee Benefits
        Range("J17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C55/1000"
        
        'External services/Freelancer
        Range("J18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C55/1000"
        
        'Legal and Consulting Costs
        Range("J19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C55/1000"
      
        'Audit Costs
        Range("J20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C55/1000"
        
     'License Costs
        Range("J21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C55/1000"
        
         'Marketing Expenses
        Range("J22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C55/1000"
        
          'Travel Expenses
        Range("J23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C55/1000"
        
          'Car Expenses
        Range("J24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C55/1000"
        
          'Office Costs
        Range("J25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C55/1000"
        
        'Hosting Costs
        Range("J26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C55/1000"
        
        'Admin/Other Expenses
        Range("J27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C55/1000"
        
        'Total Operating Expenses
        Range("J28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C55/1000"
        
        'EBITDA
        Range("J29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C55/1000"
        
         'Depreciation
        Range("J30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C55/1000"
        
         'Capitalised Costs
        Range("J31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C55/1000"
        
        'Depreciation Capitalised Costs
        Range("J32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C55/1000"
        
        'M&A Expenses
        Range("J33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C55/1000"
        
         'EBIT
        Range("J34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C55/1000"
        
         'Interest Income
        Range("J35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C55/1000"
        
        'Interest Expense
        Range("J36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C55/1000"
        
        'Extraordinary Income
        Range("J37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C55/1000"
        
         'Extraordinary Expenses
        Range("J38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C55/1000"
        
        'Earnings Before Tax
        Range("J39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C55/1000"
        
        'Tax
        Range("J40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C55/1000"
        
         'Earnings After Tax
        Range("J41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C55/1000"
        
        
        
        
        'Sep copy values from monster P&L


'Sales SaaS Licenses
        Range("K6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C56/1000"
        
 'Sales SaaS Service
        Range("K7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C56/1000"
        
    'Sales Media Posting
        Range("K8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C56/1000"
        
        'Sales Commission
        Range("K9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C56/1000"
        
        'Total Sales
        Range("K10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C56/1000"
        
        'Media Posting
        Range("K11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C56/1000"
        
        'Others (Textkernel)
        Range("K12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C56/1000"
        
        'COGS
        Range("K13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C56/1000"
        
        'Other Income/(Expense)
        Range("K14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C56/1000"
        
        'Gross Profit
        Range("K15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C56/1000"
        
         'Employee Benefits
        Range("K17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C56/1000"
        
        'External services/Freelancer
        Range("K18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C56/1000"
        
        'Legal and Consulting Costs
        Range("K19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C56/1000"
      
        'Audit Costs
        Range("K20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C56/1000"
        
     'License Costs
        Range("K21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C56/1000"
        
         'Marketing Expenses
        Range("K22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C56/1000"
        
          'Travel Expenses
        Range("K23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C56/1000"
        
          'Car Expenses
        Range("K24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C56/1000"
        
          'Office Costs
        Range("K25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C56/1000"
        
        'Hosting Costs
        Range("K26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C56/1000"
        
        'Admin/Other Expenses
        Range("K27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C56/1000"
        
        'Total Operating Expenses
        Range("K28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C56/1000"
        
        'EBITDA
        Range("K29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C56/1000"
        
         'Depreciation
        Range("K30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C56/1000"
        
         'Capitalised Costs
        Range("K31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C56/1000"
        
        'Depreciation Capitalised Costs
        Range("K32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C56/1000"
        
        'M&A Expenses
        Range("K33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C56/1000"
        
         'EBIT
        Range("K34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C56/1000"
        
         'Interest Income
        Range("K35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C56/1000"
        
        'Interest Expense
        Range("K36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C56/1000"
        
        'Extraordinary Income
        Range("K37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C56/1000"
        
         'Extraordinary Expenses
        Range("K38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C56/1000"
        
        'Earnings Before Tax
        Range("K39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C56/1000"
        
        'Tax
        Range("K40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C56/1000"
        
         'Earnings After Tax
        Range("K41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C56/1000"
        
        
        
        
        'Oct copy values from monster P&L


'Sales SaaS Licenses
        Range("L6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C57/1000"
        
 'Sales SaaS Service
        Range("L7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C57/1000"
        
    'Sales Media Posting
        Range("L8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C57/1000"
        
        'Sales Commission
        Range("L9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C57/1000"
        
        'Total Sales
        Range("L10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C57/1000"
        
        'Media Posting
        Range("L11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C57/1000"
        
        'Others (Textkernel)
        Range("L12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C57/1000"
        
        'COGS
        Range("L13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C57/1000"
        
        'Other Income/(Expense)
        Range("L14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C57/1000"
        
        'Gross Profit
        Range("L15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C57/1000"
        
         'Employee Benefits
        Range("L17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C57/1000"
        
        'External services/Freelancer
        Range("L18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C57/1000"
        
        'Legal and Consulting Costs
        Range("L19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C57/1000"
      
        'Audit Costs
        Range("L20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C57/1000"
        
     'License Costs
        Range("L21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C57/1000"
        
         'Marketing Expenses
        Range("L22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C57/1000"
        
          'Travel Expenses
        Range("L23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C57/1000"
        
          'Car Expenses
        Range("L24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C57/1000"
        
          'Office Costs
        Range("L25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C57/1000"
        
        'Hosting Costs
        Range("L26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C57/1000"
        
        'Admin/Other Expenses
        Range("L27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C57/1000"
        
        'Total Operating Expenses
        Range("L28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C57/1000"
        
        'EBITDA
        Range("L29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C57/1000"
        
         'Depreciation
        Range("L30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C57/1000"
        
         'Capitalised Costs
        Range("L31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C57/1000"
        
        'Depreciation Capitalised Costs
        Range("L32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C57/1000"
        
        'M&A Expenses
        Range("L33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C57/1000"
        
         'EBIT
        Range("L34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C57/1000"
        
         'Interest Income
        Range("L35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C57/1000"
        
        'Interest Expense
        Range("L36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C57/1000"
        
        'Extraordinary Income
        Range("L37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C57/1000"
        
         'Extraordinary Expenses
        Range("L38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C57/1000"
        
        'Earnings Before Tax
        Range("L39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C57/1000"
        
        'Tax
        Range("L40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C57/1000"
        
         'Earnings After Tax
        Range("L41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C57/1000"
        
        
        
        
        'Nov copy values from monster P&L


'Sales SaaS Licenses
        Range("M6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C58/1000"
        
 'Sales SaaS Service
        Range("M7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C58/1000"
        
    'Sales Media Posting
        Range("M8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C58/1000"
        
        'Sales Commission
        Range("M9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C58/1000"
        
        'Total Sales
        Range("M10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C58/1000"
        
        'Media Posting
        Range("M11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C58/1000"
        
        'Others (Textkernel)
        Range("M12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C58/1000"
        
        'COGS
        Range("M13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C58/1000"
        
        'Other Income/(Expense)
        Range("M14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C58/1000"
        
        'Gross Profit
        Range("M15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C58/1000"
        
         'Employee Benefits
        Range("M17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C58/1000"
        
        'External services/Freelancer
        Range("M18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C58/1000"
        
        'Legal and Consulting Costs
        Range("M19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C58/1000"
      
        'Audit Costs
        Range("M20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C58/1000"
        
     'License Costs
        Range("M21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C58/1000"
        
         'Marketing Expenses
        Range("M22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C58/1000"
        
          'Travel Expenses
        Range("M23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C58/1000"
        
          'Car Expenses
        Range("M24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C58/1000"
        
          'Office Costs
        Range("M25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C58/1000"
        
        'Hosting Costs
        Range("M26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C58/1000"
        
        'Admin/Other Expenses
        Range("M27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C58/1000"
        
        'Total Operating Expenses
        Range("M28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C58/1000"
        
        'EBITDA
        Range("M29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C58/1000"
        
         'Depreciation
        Range("M30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C58/1000"
        
         'Capitalised Costs
        Range("M31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C58/1000"
        
        'Depreciation Capitalised Costs
        Range("M32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C58/1000"
        
        'M&A Expenses
        Range("M33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C58/1000"
        
         'EBIT
        Range("M34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C58/1000"
        
         'Interest Income
        Range("M35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C58/1000"
        
        'Interest Expense
        Range("M36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C58/1000"
        
        'Extraordinary Income
        Range("M37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C58/1000"
        
         'Extraordinary Expenses
        Range("M38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C58/1000"
        
        'Earnings Before Tax
        Range("M39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C58/1000"
        
        'Tax
        Range("M40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C58/1000"
        
         'Earnings After Tax
        Range("M41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C58/1000"
        
                
                
                
        'Dez copy values from monster P&L


'Sales SaaS Licenses
        Range("N6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C59/1000"
        
 'Sales SaaS Service
        Range("N7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C59/1000"
        
    'Sales Media Posting
        Range("N8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C59/1000"
        
        'Sales Commission
        Range("N9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C59/1000"
        
        'Total Sales
        Range("N10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C59/1000"
        
        'Media Posting
        Range("N11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C59/1000"
        
        'Others (Textkernel)
        Range("N12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C59/1000"
        
        'COGS
        Range("N13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C59/1000"
        
        'Other Income/(Expense)
        Range("N14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C59/1000"
        
        'Gross Profit
        Range("N15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C59/1000"
        
         'Employee Benefits
        Range("N17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C59/1000"
        
        'External services/Freelancer
        Range("N18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C59/1000"
        
        'Legal and Consulting Costs
        Range("N19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C59/1000"
      
        'Audit Costs
        Range("N20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C59/1000"
        
     'License Costs
        Range("N21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C59/1000"
        
         'Marketing Expenses
        Range("N22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C59/1000"
        
          'Travel Expenses
        Range("N23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C59/1000"
        
          'Car Expenses
        Range("N24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C59/1000"
        
          'Office Costs
        Range("N25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C59/1000"
        
        'Hosting Costs
        Range("N26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C59/1000"
        
        'Admin/Other Expenses
        Range("N27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C59/1000"
        
        'Total Operating Expenses
        Range("N28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C59/1000"
        
        'EBITDA
        Range("N29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C59/1000"
        
         'Depreciation
        Range("N30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C59/1000"
        
         'Capitalised Costs
        Range("N31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C59/1000"
        
        'Depreciation Capitalised Costs
        Range("N32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C59/1000"
        
        'M&A Expenses
        Range("N33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C59/1000"
        
         'EBIT
        Range("N34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C59/1000"
        
         'Interest Income
        Range("N35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C59/1000"
        
        'Interest Expense
        Range("N36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C59/1000"
        
        'Extraordinary Income
        Range("N37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C59/1000"
        
         'Extraordinary Expenses
        Range("N38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C59/1000"
        
        'Earnings Before Tax
        Range("N39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C59/1000"
        
        'Tax
        Range("N40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C59/1000"
        
         'Earnings After Tax
        Range("N41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C59/1000"
        
        
        
        
        'YTD 2021 copy values
        
       
   'Sales SaaS Licenses
       Range("P6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C67/1000"
        
 'Sales SaaS Service
       Range("P7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C67/1000"
        
    'Sales Media Posting
       Range("P8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C67/1000"
        
        'Sales Commission
       Range("P9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C67/1000"
        
        'Total Sales
       Range("P10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C67/1000"
        
        'Media Posting
       Range("P11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C67/1000"
        
        'Others (Textkernel)
       Range("P12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C67/1000"
        
        'COGS
       Range("P13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C67/1000"
        
        'Other Income/(Expense)
       Range("P14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C67/1000"
        
        'Gross Profit
       Range("P15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C67/1000"
        
         'Employee Benefits
       Range("P17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C67/1000"
        
        'External services/Freelancer
       Range("P18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C67/1000"
        
        'Legal and Consulting Costs
       Range("P19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C67/1000"
      
        'Audit Costs
       Range("P20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C67/1000"
        
     'License Costs
       Range("P21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C67/1000"
        
         'Marketing Expenses
       Range("P22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C67/1000"
        
          'Travel Expenses
       Range("P23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C67/1000"
        
          'Car Expenses
       Range("P24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C67/1000"
        
          'Office Costs
       Range("P25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C67/1000"
        
        'Hosting Costs
       Range("P26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C67/1000"
        
        'Admin/Other Expenses
       Range("P27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C67/1000"
        
        'Total Operating Expenses
       Range("P28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C67/1000"
        
        'EBITDA
       Range("P29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C67/1000"
        
         'Depreciation
       Range("P30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C67/1000"
        
         'Capitalised Costs
       Range("P31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C67/1000"
        
        'Depreciation Capitalised Costs
       Range("P32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C67/1000"
        
        'M&A Expenses
       Range("P33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C67/1000"
        
         'EBIT
       Range("P34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C67/1000"
        
         'Interest Income
       Range("P35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C67/1000"
        
        'Interest Expense
       Range("P36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C67/1000"
        
        'Extraordinary Income
       Range("P37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C67/1000"
        
         'Extraordinary Expenses
       Range("P38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C67/1000"
        
        'Earnings Before Tax
       Range("P39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C67/1000"
        
        'Tax
       Range("P40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C67/1000"
        
         'Earnings After Tax
       Range("P41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C67/1000"
        
        
        
       
       
       'YTD vs. Budget
       
       
        
        
        
        
        
        'format cells
                
         Range("C6:N41").Select
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"









'YTD vs Budget


'Sales SaaS Licenses
        Range("S6").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R9C64/1000"
        
 'Sales SaaS Service
        Range("S7").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R12C64/1000"
        
    'Sales Media Posting
        Range("S8").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R16C64/1000"
        
        'Sales Commission
        Range("S9").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R18C64/1000"
        
        'Total Sales
        Range("S10").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R19C64/1000"
        
        'Media Posting
        Range("S11").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R20C64/1000"
        
        'Others (Textkernel)
        Range("S12").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R21C64/1000"
        
        'COGS
        Range("S13").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R22C64/1000"
        
        'Other Income/(Expense)
        Range("S14").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R23C64/1000"
        
        'Gross Profit
        Range("S15").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R24C64/1000"
        
         'Employee Benefits
        Range("S17").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R25C64/1000"
        
        'External services/Freelancer
        Range("S18").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R26C64/1000"
        
        'Legal and Consulting Costs
        Range("S19").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R27C64/1000"
      
        'Audit Costs
        Range("S20").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R28C64/1000"
        
     'License Costs
        Range("S21").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R29C64/1000"
        
         'Marketing Expenses
        Range("S22").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R30C64/1000"
        
          'Travel Expenses
        Range("S23").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R31C64/1000"
        
          'Car Expenses
        Range("S24").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R32C64/1000"
        
          'Office Costs
        Range("S25").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R33C64/1000"
        
        'Hosting Costs
        Range("S26").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R34C64/1000"
        
        'Admin/Other Expenses
        Range("S27").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R35C64/1000"
        
        'Total Operating Expenses
        Range("S28").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R36C64/1000"
        
        'EBITDA
        Range("S29").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R37C64/1000"
        
        'Depreciation
        Range("S30").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R38C64/1000"
        
         'Capitalised Costs
        Range("S31").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R39C64/1000"
        
        'Depreciation Capitalised Costs
        Range("S32").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R40C64/1000"
        
        'M&A Expenses
        Range("S33").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R41C64/1000"
        
         'EBIT
        Range("S34").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R42C64/1000"
        
         'Interest Income
        Range("S35").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R43C64/1000"
        
        'Interest Expense
        Range("S36").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R44C64/1000"
        
        'Extraordinary Income
        Range("S37").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R45C64/1000"
        
         'Extraordinary Expenses
        Range("S38").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R46C64/1000"
        
        'Earnings Before Tax
        Range("S39").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R47C64/1000"
        
        'Tax
        Range("S40").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R48C64/1000"
        
         'Earnings After Tax
        Range("S41").Formula = _
        "='[SG Group_P&L data base 2019-2022.xlsx]P&L Group'!R49C64/1000"
End Sub





