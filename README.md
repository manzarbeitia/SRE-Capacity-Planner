# 📊 SRE Capacity Planner CLI

[![Node.js](https://img.shields.io/badge/Node.js-18.x-green.svg)](https://nodejs.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](#contributing)

A powerful Node.js Command Line Interface (CLI) that generates highly precise, Enterprise-grade **Capacity Planning Excel Dashboards**. 

Stop guessing how many users your infrastructure can handle. This tool moves beyond simple linear math (Little's Law) and integrates real-world **Site Reliability Engineering (SRE)** principles, queueing mechanics, and the **Universal Scalability Law (USL)** to pinpoint your exact architectural bottlenecks before they cause downtime.

## ✨ Key Features

* **📈 Non-Linear Scaling Math:** Implements Neil Gunther's Universal Scalability Law (USL) to account for multithreading contention and coherency penalties.
* **🛡️ SRE Redundancy Built-in:** Calculates limits assuming N+1 or N+X hardware failures.
* **⚙️ Advanced Architecture Support:** Native support for Event Bus / Background workers CPU penalties, Redis/Memcached hit-rates, and ORM DB connection holding times.
* **🌍 Multi-Language Output:** Generates Excel spreadsheets in both English (`en`) and Spanish (`es`).
* **📊 Visual Dashboard:** Conditionally formatted DataBars to instantly spot your weakest link (CPU, RAM/Threads, Network, or Database).
* **🔒 Bulletproof Formulas:** The generated Excel is packed with `MAX(1, ...)` safeties to prevent #DIV/0! errors if fields are left blank.

## 🚀 Installation

Ensure you have [Node.js](https://nodejs.org/) installed on your machine.

1. Clone the repository:
```bash
git clone https://github.com/YOUR_USERNAME/SRE-Capacity-Planner.git
cd SRE-Capacity-Planner
```

2. Install the required dependencies:
```bash
npm install
```

## 💻 Usage

Run the CLI tool using Node. You can customize the output file name and the language of the Excel spreadsheet.

### Basic Command (Defaults to English & 'Capacity_Planning_Default.xlsx')
```bash
node index.js
```

### Custom File Name
```bash
node index.js --file_name=Production_Cluster_Review
```

### Change Language to Spanish
```bash
node index.js --lang=es --file_name=Calculadora_Infra_ES
```

### CLI Arguments
| Argument | Description | Default | Options |
| :--- | :--- | :--- | :--- |
| `--file_name=` | The name of the generated Excel file. | `Capacity_Planning_Default.xlsx` | Any valid filename string. |
| `--lang=` | The language of the spreadsheet headers, descriptions, and formulas. | `en` | `en` (English), `es` (Spanish) |

## 🏗️ How the Excel Dashboard Works

Once generated, open the Excel file. It is divided into three distinct tabs:

1.  **1. CONTROL PANEL (Inputs):** The only tab you need to edit. Enter your expected user demand, hardware specs (Cores, RAM, Network), software profiling (Latency, Payload sizes), and architecture details (DB Pool, Cache ratio).
2.  **2. EXECUTIVE DASHBOARD:** A read-only visual dashboard. It automatically calculates the exact maximum RPS (Requests Per Second) your system can handle, flags your critical bottleneck, and tells you exactly how many simultaneous users your current setup can safely support.
3.  **3. MATH ENGINE (Hidden):** The core where the SRE math happens. It calculates OS Page Cache reservations, Event Bus CPU penalties, and USL degradation.

## 🧠 The Math Behind It

Traditional capacity calculators assume that if 1 CPU core handles 100 users, 16 cores will handle 1,600 users. **This is mathematically false in the real world.** This tool uses the **Universal Scalability Law** to plot a realistic curve, acknowledging that adding concurrent threads eventually degrades performance due to resource locks (Contention) and state synchronization (Coherency). Furthermore, it isolates purely computational CPU time from total connection time, ensuring RAM and Thread Pool limits are calculated accurately.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/YOUR_USERNAME/SRE-Capacity-Planner/issues).

## 📝 License

This project is [MIT](https://opensource.org/licenses/MIT) licensed.
