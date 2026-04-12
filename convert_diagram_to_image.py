"""
Convert Mermaid diagram to PNG image
"""
import subprocess
import sys

# First, try to install mermaid-cli via npm
try:
    print("Installing mermaid-cli...")
    subprocess.run(["npm", "install", "-g", "mermaid-cli"], check=False, capture_output=True)
    print("✅ mermaid-cli installed")
except:
    print("⚠️ npm not found, trying alternative method...")

# Save the Mermaid diagram to a file
mermaid_content = """flowchart TD
  subgraph BankSources [Banking Data Sources]
    CoreBank["💳 Core Banking System\n(Accounts, Balances)"]
    TxnLog["📊 Transaction Log\n(Deposits, Withdrawals, Transfers)"]
    CRMData["👤 CRM System\n(Customer KYC, Profile)"]
    LoanSys["📋 Loan Management\n(Applications, Repayments)"]
  end

  subgraph GovLayer [Governance: Banking Policies & Compliance]
    Policies["Policies: PCI-DSS, AML/KYC, GDPR\nRetention: 7 years for transactions\nClassification: Public / Internal / Confidential / PII"]
    SLAs["SLAs:\n- Account balance <5min latency\n- KYC updates <1 day\n- Fraud alerts <10 sec"]
    DataOwners["Owners:\n- Finance: Accounts, Balances\n- Compliance: KYC, AML transactions\n- Risk: Loans, Credit scores"]
  end

  subgraph Ingestion [ETL/Ingest Pipeline]
    E1["Ingest & Normalize\n(PII masking stage)"]
    E2["Deduplication & Allocation\n(Customer matching)"]
    E3["Store in Warehouse\n(Snowflake or BigQuery)"]
  end

  subgraph DataWarehouse [Data Warehouse - Banking Entities]
    Customers["Customers Table\n(ID, Name, KYC_Status, Risk_Level)"]
    Accounts["Accounts Table\n(Account_ID, Customer_ID, Balance, Type)"]
    Transactions["Transactions Table\n(Txn_ID, Account_ID, Amount, Date, Type)"]
    Loans["Loans Table\n(Loan_ID, Customer_ID, Amount, Status)"]
  end

  subgraph QualityChecks [Data Quality Rules & Monitoring]
    R1["Customers:\n✓ No null KYC\n✓ Email valid\n✓ Unique SSN"]
    R2["Accounts:\n✓ Balance >= 0\n✓ Matched to Customer\n✓ No orphans"]
    R3["Transactions:\n✓ Amount > 0\n✓ Txn_Date <= Today\n✓ Account exists\n✓ Reconcile to GL"]
    R4["Loans:\n✓ Loan_Amount > 0\n✓ Start_Date <= Today\n✓ Customer verified"]
    
    Monitor["📈 Monitor & Alert\n- Daily transaction count trend\n- Failed reconciliation %\n- Duplicate customer accounts\n- Negative balance flags"]
  end

  subgraph Remediation [Data Remediation]
    Auto["Automated Fixes:\n- Trim/standardize names\n- Dedup by SSN + DOB\n- Backfill missing dates"]
    Manual["Manual Review:\n- Conflicting customer records\n- Suspicious transactions\n- High-risk customers"]
    Incident["Incident Tracking:\n- Create ticket\n- Track MTTR\n- RCA & prevention"]
  end

  subgraph Compliance [Compliance & Audit]
    AML["AML Checks:\n- Sanction screening\n- Transaction monitoring\n- Suspicious activity reports"]
    PII["PII Protection:\n- Mask SSN/Account# for analysts\n- Encrypt at rest\n- Access logs"]
    Audit["Audit Trail:\n- Who accessed what\n- When & for what purpose\n- Data lineage"]
  end

  subgraph Consumption [Analytics & Consumer Applications]
    FinReport["📄 Financial Reports\n(Balance sheets, P&L)"]
    RiskDash["⚠️ Risk Dashboard\n(Credit exposure, NPL %)"]
    FraudDetect["🚨 Fraud Detection\n(Anomaly ML model)"]
    CustSeg["🎯 Customer Segmentation\n(Lifetime value, churn risk)"]
  end

  subgraph Feedback [Consumer Feedback Loop]
    Issues["Report Issues:\n- Wrong balance in report\n- Missed fraud\n- Late transactions"]
    UpdateSLAs["Update SLAs & Rules\n- Tighten thresholds\n- Add new checks"]
  end

  CoreBank --> E1
  TxnLog --> E1
  CRMData --> E1
  LoanSys --> E1
  
  E1 --> E2
  E2 --> E3
  
  E3 --> Customers
  E3 --> Accounts
  E3 --> Transactions
  E3 --> Loans
  
  Customers --> R1
  Accounts --> R2
  Transactions --> R3
  Loans --> R4
  
  R1 --> Monitor
  R2 --> Monitor
  R3 --> Monitor
  R4 --> Monitor
  
  Monitor --> Auto
  Auto --> Incident
  Monitor --> Manual
  Manual --> Incident
  
  Incident --> E3
  
  Transactions --> AML
  Customers --> PII
  Transactions --> Audit
  
  Customers --> FinReport
  Accounts --> FinReport
  Transactions --> FinReport
  
  Transactions --> RiskDash
  Accounts --> RiskDash
  Loans --> RiskDash
  
  Transactions --> FraudDetect
  Customers --> CustSeg
  Accounts --> CustSeg
  
  FinReport --> Issues
  RiskDash --> Issues
  FraudDetect --> Issues
  CustSeg --> Issues
  
  Issues --> UpdateSLAs
  UpdateSLAs --> GovLayer
  
  GovLayer --> Ingestion
  GovLayer --> QualityChecks
  
  style BankSources fill:#e1f5ff,stroke:#01579b,stroke-width:2px
  style GovLayer fill:#f3e5f5,stroke:#4a148c,stroke-width:2px
  style Ingestion fill:#e8f5e9,stroke:#1b5e20,stroke-width:2px
  style DataWarehouse fill:#fce4ec,stroke:#880e4f,stroke-width:2px
  style QualityChecks fill:#fff3e0,stroke:#e65100,stroke-width:2px
  style Remediation fill:#f1f8e9,stroke:#33691e,stroke-width:2px
  style Compliance fill:#ede7f6,stroke:#311b92,stroke-width:2px
  style Consumption fill:#e0f2f1,stroke:#004d40,stroke-width:2px
  style Feedback fill:#fff9c4,stroke:#f57f17,stroke-width:2px
"""

# Save mermaid file
mermaid_file = r"e:\VScode26\banking_diagram.mmd"
with open(mermaid_file, 'w') as f:
    f.write(mermaid_content)
print(f"✅ Mermaid diagram saved: {mermaid_file}")

# Try to convert using mmdc
try:
    output_png = r"e:\VScode26\Banking_Data_Flow_Diagram.png"
    result = subprocess.run(
        ["mmdc", "-i", mermaid_file, "-o", output_png],
        capture_output=True,
        text=True,
        timeout=30
    )
    if result.returncode == 0:
        print(f"✅ PNG created: {output_png}")
    else:
        print(f"⚠️ mmdc conversion failed: {result.stderr}")
        raise Exception("mmdc failed")
except FileNotFoundError:
    print("⚠️ mmdc not found, trying Playwright method...")
    try:
        from playwright.sync_api import sync_playwright
        
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 1920, "height": 1080})
            
            # Create HTML with Mermaid
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <script src="https://cdn.jsdelivr.net/npm/mermaid@latest/dist/mermaid.min.js"></script>
                <style>
                    body {{ margin: 0; padding: 20px; background: white; }}
                    .mermaid {{ display: flex; justify-content: center; }}
                </style>
            </head>
            <body>
                <div class="mermaid">
{mermaid_content}
                </div>
                <script>
                    mermaid.contentLoaderFunction = () => {{
                        return fetch('data:text/mermaid;charset=utf-8,' + encodeURIComponent(`{mermaid_content}`)).then(d => d.text());
                    }};
                    mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
                    mermaid.contentLoaderFunction = undefined;
                    mermaid.run();
                </script>
            </body>
            </html>
            """
            
            # Save HTML
            html_file = r"e:\VScode26\banking_diagram_temp.html"
            with open(html_file, 'w') as f:
                f.write(html_content)
            
            # Navigate and screenshot
            page.goto(f"file:///{html_file.replace(chr(92), '/')}")
            page.wait_for_selector(".mermaid svg", timeout=10000)
            
            output_png = r"e:\VScode26\Banking_Data_Flow_Diagram.png"
            page.screenshot(path=output_png, full_page=True)
            browser.close()
            
            print(f"✅ PNG created using Playwright: {output_png}")
    except Exception as e:
        print(f"❌ Playwright method failed: {e}")
        print("Please install mermaid-cli: npm install -g mermaid-cli")
        sys.exit(1)

print("\n✅ Done! Image file created successfully.")
