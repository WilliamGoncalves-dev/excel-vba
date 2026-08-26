# Service Billing & Tax Automation

An Excel/VBA-based business process automation solution designed to streamline the preparation, validation, and operational processing of service invoices across multiple branches in Brazil.

The solution was originally developed in a professional billing environment where service invoices had to be prepared for different branches, customers, contracts, municipalities, and service types.

The workbook consolidated data from ERP exports, applied business rules through Excel formulas, automated repetitive tasks with VBA, generated Outlook email drafts, and maintained an operational billing log.

---

## 📊 Business Impact

The original process required approximately **30 minutes per sales order** to manually review, calculate, prepare, and organize the information required for invoice processing.

After implementing the solution, the same workflow could be completed in **up to approximately 5 minutes**.

| Metric                          |             Before |                After |
| ------------------------------- | -----------------: | -------------------: |
| Processing time per sales order |            ~30 min |              ≤ 5 min |
| Time reduction                  |                  — |             **~83%** |
| Main process                    |             Manual | Automated / Assisted |
| Data retrieval                  |      Manual lookup |     Automated lookup |
| Invoice description             |       Manual entry |       Formula-driven |
| Tax logic                       | Manual calculation |           Rule-based |
| Email preparation               |             Manual |         VBA-assisted |
| Billing tracking                |            Limited |    Automated summary |

The approximately 83% reduction is calculated from the reported change from 30 minutes to 5 minutes per sales order.

---

## 👤 Role & Contribution

I worked as a **Billing Specialist**, responsible for the issuance and processing of service invoices.

While performing this role, I identified repetitive and manual steps in the billing workflow and independently designed and implemented an Excel/VBA solution to standardize and automate part of the process.

My contribution included:

- Process and business rule analysis
- Excel workbook and formula development
- ERP data consolidation
- VBA automation and Outlook integration
- Operational tracking and reporting
- Testing and process improvement

## 🧠 Business Context

The company operated with multiple branches across Brazil and provided different types of services.

The billing process involved:

* Multiple branches and service locations.
* Different customers and contracts.
* Different service types.
* Different municipalities.
* Service-specific descriptions.
* Payment terms.
* ISS-related rules.
* Manual interaction with the municipal invoice portal.
* ERP data that was not directly integrated with the municipal systems.

The ERP contained the operational information required for billing, but it did not provide a direct API integration with the municipal invoice platforms.

As a result, the billing process depended heavily on manual data retrieval, calculations, validation, and transcription.

The solution was created to introduce an intermediate automation layer using Excel.

---

## 🎯 Problem

The original workflow required the user to repeatedly:

1. Find the correct Sales Order.
2. Retrieve the corresponding contract.
3. Identify the customer and responsible branch.
4. Determine the service being billed.
5. Prepare the invoice description.
6. Normalize address information.
7. Determine payment terms and due dates.
8. Calculate the applicable tax bases.
9. Determine ISS-related values.
10. Determine the appropriate operation code based on location rules.
11. Prepare the customer email.
12. Identify the correct recipients.
13. Record the processed billing operation.
14. Export operational reports.

Many of these steps involved repetitive lookups and manual transcription.

The objective was therefore not simply to "make an Excel spreadsheet", but to **standardize and automate the workflow while preserving the existing business rules**.

---

# 🏗️ Solution Architecture

The workbook acts as an automation and decision-support layer between ERP-generated data and the billing workflow.

```text
                    ERP
                     │
        ┌────────────┼─────────────┐
        │            │             │
        ▼            ▼             ▼
  ORDEM DE VENDA  CLIENTES       ISS
        │            │             │
        └────────────┼─────────────┘
                     │
                     ▼
          TRIBUTAÇÃO FORA
          DO MUNICÍPIO
                     │
          ┌──────────┴──────────┐
          │                     │
          ▼                     ▼
      SERVIÇOS               E-MAIL
          │                     │
          └──────────┬──────────┘
                     ▼
              Business Rules
                     │
                     ▼
             Invoice Preparation
                     │
                     ▼
              Municipal Portal
                     │
                     ▼
              Outlook / Customer
                     │
                     ▼
                  RESUMO
                     │
                     ▼
              Operational Reports
```

---

# 📑 Workbook Structure

The workbook is composed of several interconnected worksheets.

## `Tributação fora do Municipio`

The main operational worksheet.

It acts as the central interface where the user enters or selects the Sales Order and obtains the information required for invoice preparation.

The worksheet consolidates information from the supporting datasets and applies the business rules required to prepare the invoice.

Examples include:

* Service description
* Customer information
* Contract
* Branch/site
* Address
* Payment terms
* Due date
* Tax calculations
* ISS rate
* Operation code
* Net invoice amount

---

## `ORDEM DE VENDA`

ERP-generated source data containing information related to the Sales Orders.

Relevant fields include:

* Sales Order number
* Gross value
* Total service value
* Contract number
* Sales/service description
* Customer
* Branch/site
* Address
* State
* Municipality

The municipality and state fields were added to the ERP dataset through a customization requested from the IT team because these attributes were required by the automation logic.

---

## `Clientes`

Customer and contract-related information used to retrieve:

* Contract
* Payment terms
* Customer
* Customer ID
* Responsible branch/site

The contract number retrieved from the Sales Order acts as an important lookup key.

---

## `ISS`

A reference table containing municipality-level ISS rates used by the workbook when calculating invoice-related tax information.

The main worksheet uses nested lookup logic to retrieve the appropriate rate.

---

## `SERVIÇOS`

A manually created parameter table.

The ERP provided long service descriptions as text. Instead of repeatedly typing or selecting those descriptions manually, internal numeric identifiers were created.

The relationship is approximately:

```text
Internal Service Code
        │
        ▼
SERVIÇOS
        │
        ▼
Standardized Service Description
        │
        ▼
Invoice Description
```

This reduced repetitive typing and standardized the service descriptions used during invoice preparation.

---

## `E-MAIL`

A dataset containing customer contact information.

The main challenge was that the same contact could appear under multiple customer records or contracts.

For example:

```text
Customer / Contract A → john@example.com
Customer / Contract B → john@example.com
Customer / Contract C → finance@example.com
```

The VBA solution therefore builds a unique recipient list before creating the email draft.

---

## `RESUMO`

An operational log generated by VBA.

Each processed invoice adds a new record containing information such as:

* Sales Order
* Contract
* Customer
* Invoice value
* Net value
* Operation code
* Branch
* Processing date

This transformed the workbook from a simple calculation tool into a lightweight operational tracking system.

---

# 🧮 Business Rules

A significant portion of the automation was implemented using Excel formulas.

The formulas connect the datasets and encode the business rules required by the billing process.

### Service Description

The main worksheet uses lookup logic to transform an internal service code into a standardized description and combine it with information retrieved from the Sales Order.

This allowed invoice descriptions to be generated automatically rather than manually typed for every invoice.

---

### Address Normalization

Address information exported from the ERP could contain line breaks.

The workbook uses text functions such as `SUBSTITUIR`, `EXT.TEXTO`, `CARACT(10)` and `CARACT(13)` to normalize the value into a single-line address.

---

### Payment Terms and Due Dates

Payment conditions retrieved from the customer dataset are used to determine the invoice/billing due date.

The logic supports different business scenarios, including:

* Fixed payment dates
* Seven-day terms
* Multiple installments such as `30/60`
* Numeric payment periods

This means the due date is generated from the underlying customer/payment configuration instead of being manually calculated for each invoice.

---

### Tax Calculation

The workbook calculates the values required for invoice validation based on the configured business rules.

The calculation includes elements such as:

* Gross invoice amount
* Equipment/service allocation
* INSS
* ISS
* Tax bases
* Net invoice amount

The resulting values were used to validate whether the invoice issued through the municipal system matched the expected values calculated by the workbook.

---

# 🧭 Operation Code Logic

One of the most complex parts of the workbook was the determination of the operation code.

The calculation considers three geographic variables:

```text
Service Provider
       │
       ├──────────────┐
       │              │
       ▼              ▼
   Municipality    Municipality
   of Provider      of Service
                       │
                       ▼
                 Municipality
                   of Customer
```

The combinations between these locations determine which configured operation code should be used.

The original workbook used a nested `SE` structure with multiple `PROCV` operations to compare the relevant municipalities.

The configured codes were:

| Code     | Business scenario                                                                                                                             |
| -------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| **2015** | Service taxed within the municipality, with ISS responsibility assigned to the customer according to the company's configured rule            |
| **2024** | ISS retention handled by the service provider according to the company's configured rule                                                      |
| **2031** | Service performed outside the municipality of origin of both provider and customer, normally with ISS responsibility assigned to the customer |

These codes represent the company's operational configuration at the time the solution was developed.

> **Important:** This documentation describes the business logic implemented in the original workbook. It is not intended to provide current Brazilian tax or legal guidance.

---

# 🤖 VBA Automation

After the invoice was issued through the municipal portal and the workbook was fully populated, VBA automated the next operational steps.

---

## 📧 Automated Email Preparation

The `EnviarEmailSemtelefone` macro:

1. Retrieves the customer identifier from the main worksheet.
2. Searches the `E-MAIL` worksheet.
3. Collects all matching email addresses.
4. Uses `Scripting.Dictionary` to eliminate duplicate addresses.
5. Creates an Outlook instance.
6. Generates an email draft.
7. Builds the subject using:

   * Sales Order
   * Contract
   * Customer
8. Displays the message for human review.

The current implementation intentionally uses:

```vb
.Display
```

instead of:

```vb
.Send
```

This keeps a human approval step before the message is sent.

---

## 🧹 Duplicate Email Handling

A particularly relevant part of the automation was the use of a dictionary:

```vb
Set emailDict = CreateObject("Scripting.Dictionary")
```

The dictionary stores email addresses already encountered.

Before adding an address:

```vb
If Not emailDict.Exists(email) Then
```

the macro checks whether it has already been collected.

This prevents the same contact from receiving duplicate recipient entries when that person is associated with multiple customer records or contracts.

---

# 📋 Automated Operational Log

The email macro also records the processed invoice in the `RESUMO` worksheet.

The workflow therefore becomes:

```text
Invoice Processed
       │
       ▼
Click "Enviar e-mail"
       │
       ├── Find recipients
       ├── Remove duplicates
       ├── Create Outlook draft
       │
       ▼
Record operation
       │
       ├── Sales Order
       ├── Contract
       ├── Customer
       ├── Values
       ├── Branch
       └── Processing date
```

This was particularly useful for management because it created a centralized operational history of processed billing activities.

---

# 📤 Report Export

The workbook also contains VBA procedures for exporting operational information.

### `ExportarOrdemDeVendaSemNOMECOMDATA`

Exports the `ORDEM DE VENDA` worksheet into a separate Excel workbook and generates a timestamped filename.

### `ExportarResumo`

Exports the `RESUMO` worksheet into a monthly directory structure.

This created a separate historical reporting layer outside the main workbook.

---

# 🔄 End-to-End Workflow

The complete process can be represented as:

```text
1. ERP Data Extraction
        │
        ▼
2. Load Source Worksheets
        │
        ├── ORDEM DE VENDA
        ├── CLIENTES
        ├── ISS
        └── E-MAIL
        │
        ▼
3. Select Sales Order
        │
        ▼
4. Retrieve Contract & Customer
        │
        ▼
5. Retrieve Service Configuration
        │
        ▼
6. Generate Invoice Description
        │
        ▼
7. Normalize Address
        │
        ▼
8. Calculate Payment Due Date
        │
        ▼
9. Calculate Tax Values
        │
        ▼
10. Determine Operation Code
        │
        ▼
11. Validate / Issue Invoice
        │
        ▼
12. Generate Customer Email
        │
        ├── Retrieve recipients
        ├── Remove duplicates
        └── Create Outlook draft
        │
        ▼
13. Register Processed Invoice
        │
        ▼
14. Export Operational Reports
```

---

# 🛠️ Technologies

* **Microsoft Excel**
* **VBA**
* **Excel formulas**
* `PROCV`
* `SE`
* `SEERRO`
* `SUBSTITUIR`
* `EXT.TEXTO`
* `NÚM.CARACT`
* `Scripting.Dictionary`
* **Microsoft Outlook Automation**
* Windows file system operations

---

# 💡 Key Engineering Concepts Demonstrated

Although implemented in Excel, the project applies concepts commonly found in software and data applications:

### Data Integration

Multiple ERP-generated datasets are connected through common business identifiers such as:

* Sales Order
* Contract
* Customer ID
* Municipality
* Service Code

### Business Rule Encoding

Operational decisions that previously required manual interpretation were converted into deterministic Excel logic.

### Data Normalization

Raw ERP information was transformed into structured values suitable for invoice preparation.

### Parameterization

The `SERVIÇOS` worksheet separated service configuration from the main processing logic.

### Automation

VBA eliminated repetitive actions involving:

* Email preparation
* Recipient lookup
* Duplicate removal
* Record creation
* Report export

### Operational Traceability

The `RESUMO` worksheet created a basic audit trail of processed billing operations.

---

# 📈 Results

The most significant result was the reduction in processing time.

```text
Before
~30 minutes / Sales Order
        │
        ▼
Manual lookup + calculation + preparation
        │
        ▼
After
≤5 minutes / Sales Order
```

This represents an approximate:

**83% reduction in processing time.**

Beyond speed, the solution also improved:

* Standardization
* Repeatability
* Data retrieval
* Calculation consistency
* Email preparation
* Operational visibility
* Reporting

---

# ⚠️ Limitations

The original solution was designed around a specific operational environment and therefore contains several environment-specific dependencies.

Examples include:

* ERP export formats
* Worksheet names
* Cell references
* Customer datasets
* Municipal configurations
* Outlook availability
* Windows directory structure
* Company-specific business rules

The workbook should therefore be viewed as a **case study of process automation**, rather than a plug-and-play billing application.

---

# 🔒 Data Privacy

This portfolio version uses fictional or anonymized data.

Real customer names, email addresses, contracts, and confidential company information have been removed or replaced.

The workbook is provided strictly for portfolio and educational demonstration purposes.

---

# 📁 Repository Structure

```text
service-billing-tax-automation/
│
├── README.md
│
├── workbook/
│   └── Faturamento_Template.xlsm
│
├── vba/
│   ├── EnviarEmail.bas
│   ├── ExportarOrdemDeVenda.bas
│   └── ExportarResumo.bas
│
├── documentation/
│   ├── business-rules.md
│   └── workflow.md
│
└── screenshots/
    ├── main-sheet.png
    ├── summary.png
    └── email-draft.png
```

---

# ▶️ How to Explore the Project

The workbook is provided as a demonstration artifact using fictionalized data.

To inspect the implementation:

1. Open the `.xlsm` workbook in Microsoft Excel.
2. Review the `Tributação fora do Municipio` worksheet.
3. Inspect the supporting worksheets.
4. Open the VBA editor with `Alt + F11`.
5. Review the VBA modules.
6. Follow the formulas from the main worksheet back to their source datasets.
7. Review the `RESUMO` worksheet to understand the operational logging mechanism.

Because the original workflow depended on an ERP and a municipal invoice platform, the portfolio version should be understood as a **reproducible demonstration of the Excel/VBA automation architecture**, not a live production billing environment.

---

# 📚 What This Project Demonstrates

This project represents an example of using Excel as an automation platform rather than simply as a calculation tool.

The solution combines:

**Data → Rules → Calculations → Automation → Communication → Reporting**

That combination was the core objective of the project.

---

## 👤 Author

**William Gonçalves**

Information Systems student interested in:

* Data Analysis
* Business Intelligence
* SQL
* Python
* Java
* Process Automation
* Excel/VBA


---

## ⚠️ Disclaimer

This project is presented for educational and portfolio purposes.

The tax-related rules documented here reflect the business configuration implemented in the original solution and should not be interpreted as current legal, accounting, or tax advice.
