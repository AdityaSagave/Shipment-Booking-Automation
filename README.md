# ✈️ Shipment Booking Automation - Power Automate Desktop

## 📌 Overview

This project is a **no-code/low-code automation solution** built with **Microsoft Power Automate Desktop** that transforms a repetitive, high-pressure shipping task into a **smooth, error-free workflow**.

Designed for **logistics teams**, **customer service reps**, and **compliance-driven operations**, this automation saves time, ensures accuracy, and integrates seamlessly with daily tools like **Excel, Word, and Outlook**.

---

## ✨ Features

### 📅 1. Automated Booking Form Preparation
- 🕒 Retrieves the **current date** and inserts **day**, **month**, and **year** into an Excel template.
- 🚚 Prompts for **shipment weight** and **dimensions**, and writes them into the form.
- 🔄 Supports **multiple shipments** based on user input.

---

### 📄 2. Smart Document Editing (Word)
- 🧾 **Customs Invoice (CI)**:
  - Updates shipment **date** and **total weight**.
- 📋 **Consignment Security Declaration (CSD)**:
  - Allows user to select booking agent (from dropdown).
  - Automatically inserts the correct **ASIC number**, **signature image**, and **current date**.
  - Generates and increments **tamper tape numbers**.

---

### 📧 3. Email Automation with Outlook
- 📤 Sends:
  - Initial email with **booking form**.
  - Secondary email with **template CI & CSD** (pre-Job ID).
- 📥 Monitors inbox:
  - Extracts **Job ID** and **AWB number** from replies using **regex**.
  - Updates documents with actual values and sends final versions.

---

### 🖨️ 4. Print-Ready Dispatch
- Auto-prints final **CI** and **CSD** so physical documents are ready for the driver’s signature.

---

## 💡 Why It Matters

> 🧠 "This used to take 30–45 minutes and multiple people. Now, it's done in under **5 minutes** with full accuracy."

### ✅ Business Benefits
- **Zero manual errors** in mission-critical shipment documents.
- **Time savings** for customer service during high-volume operations.
- Fully **compliant**, traceable, and scalable automation.

---

## 🛠️ Tech Stack

| Tool/Tech | Purpose |
|----------|---------|
| **Power Automate Desktop** | Core workflow automation |
| **Microsoft Excel** | Shipment form data entry |
| **Microsoft Word** | CI & CSD document editing |
| **Outlook Desktop** | Automated emails & monitoring |
| **Regex Matching** | Parsing Job ID & AWB |
| **File/Date/Loop Controls** | Flow logic & iteration |

---

## 🚀 Getting Started

> ⚠️ This project is designed for internal enterprise use. The following folders/files must exist with proper formatting:

- `Booking form-offline/ORDER BOOKING FORM INTLOPS [date].xlsx`
- `CSD-offline/Consignment Security Declaration - [date].docx`
- `CI-offline/LDR Inc Customs Invoice [date].docx`
- Signature images for booking agents in `CSD-offline/`

**Pre-requisites:**
- Microsoft Power Automate Desktop
- Office apps installed (Excel, Word, Outlook)
- User profile paths correctly configured

---

## 🤝 Project Role

Created and deployed by **Aditya Sagave** as part of internal process automation for **Landauer Australasia**. Fully owned and maintained from design to deployment, including:

- Flowchart design
- UI dialogs
- Dynamic document edits
- Email automation
- Regex parsing for IDs

---

## 🎯 Outcome

This solution demonstrates how **no-code tools** can solve real-world business problems in logistics, with impact measurable in **time, accuracy, and compliance**.

---

## 👀 Demo Screenshot (Optional)

> *screenshot coming soon*

---

## 📫 Contact

**Aditya Sagave**  
Data Science graduate – * Macquarie University*  
✉️ [adityasagave@gmail.com](mailto: adityasagave@gmail.com)  
📞 +61 410 806 258

---

> ⭐ *If you're a recruiter or hiring manager looking for someone who blends business logic, automation, and real-world delivery: this is what I do.*

