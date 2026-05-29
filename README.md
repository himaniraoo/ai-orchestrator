# DocNexus AI — Multi-Agent Pharma Intelligence Platform

DocNexus AI is an LLM-powered multi-agent orchestration platform for pharmaceutical intelligence workflows.

The system accepts natural language healthcare analytics queries, routes requests to specialized AI agents, executes structured analysis workflows, and generates real downloadable business artifacts including:

- PowerPoint presentations (`.pptx`)
- Excel workbooks (`.xlsx`)
- Market access reports (`.docx`)
- Sandbox-executed Python analysis with charts

Built as part of the DocNexus AI Engineer Intern Assignment.

---

# Demo Capabilities

Users can submit natural language requests such as:

```json
{
  "query": "Give me a slide deck and an Excel breakdown of high-volume NSCLC oncologists in California and New York",
  "preferences": {}
}
```

or:

```json
{
  "query": "Run an analysis and show me which states have the highest concentration of high-volume NSCLC prescribers",
  "preferences": {}
}
```

# WORKING DEMO LINK

https://drive.google.com/file/d/1CMsECEgOD0tfApUqvgIVDvoamposJ1SN/view?usp=sharing

The orchestrator automatically:

1. Understands the user intent
2. Retrieves relevant physician data
3. Routes requests to specialized agents
4. Streams live orchestration updates via SSE
5. Returns downloadable artifacts and inline outputs

---

# Core Features

## Multi-Agent Orchestration

A custom orchestration loop built using Gemini native function calling.

The orchestrator:

- Interprets user queries
- Maps semantic intent to structured filters
- Calls tools dynamically
- Delegates tasks to specialized agents
- Coordinates multi-agent workflows
- Streams live execution traces to the frontend

---

# Specialized Agents

## PPT Agent

Generates PowerPoint slide decks using `python-pptx`.

Capabilities:
- Executive summary slides
- Physician rankings
- Geographic analysis
- Market access insights

Output: `.pptx`

---

## Excel Agent

Generates structured Excel workbooks using `openpyxl`.

Includes:
- Raw physician data sheet
- State × specialty pivot analysis
- ICD-10 claim volume breakdowns

Output: `.xlsx`

---

## Report Agent

Generates structured market access reports.

Features:
- Markdown report generation
- Downloadable `.docx`
- Regional physician landscape analysis

Output:
- Inline rendered markdown
- `.docx`

---

## Sandbox Agent

Generates and executes Python code dynamically inside a secure sandbox using `e2b-code-interpreter`.

Capabilities:
- Automatic code generation using Gemini
- Chart creation
- Python execution
- Error handling + self-correction retry
- Inline chart rendering

Output:
- Generated Python code
- Execution output
- Inline charts

---

# Live Orchestration Streaming (SSE)

The frontend receives live updates from the orchestrator using Server-Sent Events (SSE).

As each agent executes, the UI streams:

- Current tool execution
- Agent progress
- Artifact creation
- Sandbox execution status
- Final summaries

This provides real-time orchestration visibility instead of waiting for a single synchronous response.

---

# Preference Context System

A major architectural focus of the project was converting structured user preferences into semantically meaningful LLM context.

Example mappings:

- `NSCLC` → `C341`, `C342`
- `oncologists` → `Medical Oncology`
- `high volume` → `volume_threshold = high`
- `Northeast` → `[NY, MA, CT, NJ, PA, MD]`

The orchestrator combines:
- natural language intent
- structured preference context
- domain-specific semantic mappings

This improves:
- routing precision
- data grounding
- artifact quality
- reduction of hallucinations

---

# Architecture

```text
Frontend (React + Vite)
        ↓
FastAPI Backend
        ↓
Gemini Orchestrator Agent
        ↓
Function Calling Tool Loop
        ↓
Specialized Agents
 ├── PPT Agent
 ├── Excel Agent
 ├── Report Agent
 └── Sandbox Agent
        ↓
Artifacts + Streamed Results
```

---

# Tech Stack

## Backend

- Python
- FastAPI
- Gemini 2.5 Flash (`google-genai`)
- python-pptx
- openpyxl
- python-docx
- e2b-code-interpreter

## Frontend

- React
- Vite
- Server-Sent Events (SSE)

---

# LLM Framework Choice

The project uses Google's Gemini 2.5 Flash model through the official `google-genai` SDK.

Reasons for this choice:

- Native structured function calling support
- Fast inference speed for orchestration workflows
- Strong tool-routing performance
- Clean integration with custom orchestration loops

A custom orchestration loop was implemented instead of using higher-level frameworks like LangChain in order to better understand and control the underlying agent execution mechanics.

---

# Project Structure

```text
docnexus/
├── backend/
│   ├── main.py
│   ├── agents/
│   │   ├── orchestrator.py
│   │   ├── ppt_agent.py
│   │   ├── excel_agent.py
│   │   ├── report_agent.py
│   │   └── sandbox_agent.py
│   ├── tools/
│   │   └── physician_data.py
│   ├── data/
│   │   └── physicians.json
│   ├── artifacts/
│   ├── requirements.txt
│   └── .env
├── frontend/
└── README.md
```

---

# API Endpoints

## Health Check

```http
GET /health
```

---

## Physician Filtering Endpoint

```http
GET /physicians
```

Query parameters:
- `specialty`
- `states`
- `icd10_codes`
- `volume_threshold`

---

## Artifact Download

```http
GET /artifacts/{artifact_id}
```

---

## Query Endpoint

```http
POST /query
```

---

# Environment Variables

Create `.env` inside `backend/`

```env
GEMINI_API_KEY=your_key_here
E2B_API_KEY=your_key_here
```

Also create:

```text
backend/.env.example
```

with the same variables.

---

# Setup Instructions

## 1. Clone Repository

```bash
git clone <repo-url>
cd docnexus
```

---

## 2. Backend Setup

```bash
cd backend
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## 3. Run Backend

```bash
uvicorn main:app --reload
```

Backend runs on:

```text
http://localhost:8000
```

---

## 4. Frontend Setup

```bash
cd frontend
npm install
npm run dev
```

Frontend runs on:

```text
http://localhost:5173
```

---

# Example Queries

## Multi-Agent Query

```json
{
  "query": "Give me a slide deck and an Excel breakdown of high-volume NSCLC oncologists in California and New York",
  "preferences": {}
}
```

---

## Report Query

```json
{
  "query": "Write a two-page market access report on NSCLC physician density in the Northeast",
  "preferences": {
    "states": ["NY", "MA", "CT", "NJ", "PA", "MD"],
    "icd10_codes": ["C341", "C342"]
  }
}
```

---

## Sandbox Query

```json
{
  "query": "Run an analysis and show me which states have the highest concentration of high-volume NSCLC prescribers",
  "preferences": {}
}
```

---

# What I Would Build Next

Given more time, the next improvements would include:

- Persistent conversation history
- Real physician data API integrations
- Improved frontend analytics visualizations
- Authentication and user workspaces
- More advanced orchestration monitoring

---

# Known Limitations

- Uses mock physician data
- No persistent database layer
- Frontend optimized primarily for desktop workflows
- Limited long-running orchestration management

---

# Assignment Goals Satisfied

1. Multi-agent orchestration system  
2. Natural language query routing  
3. Specialized AI agents  
4. Real downloadable artifacts  
5. Sandbox code execution  
6. Self-correction retry loop  
7. Live orchestration streaming  
8. Structured preference context system  
9. Frontend visualization and rendering  
10. End-to-end AI workflow demonstration

