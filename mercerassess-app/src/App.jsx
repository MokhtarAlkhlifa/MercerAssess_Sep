import * as XLSX from "xlsx";
import QRCode from "qrcode";
import PptxGenJS from "pptxgenjs";
import React, { useState, useEffect, useId, createContext, useContext } from "react";
import {
  Home, User, ShieldCheck, Bell, LogOut, GraduationCap, Download,
  ChevronRight, ChevronDown, Sparkles, FileText, Users, SlidersHorizontal,
  Upload, BarChart3, Award, Search, Plus, Pencil, Trash2, Building2, Share2,
  PieChart as PieChartIcon, UploadCloud, Filter, MessageSquare, TrendingUp,
  Percent, RefreshCw, Link2, CheckCircle2, XCircle, Tag, Save, Clock,
  FileSpreadsheet, FileType, Image as ImageIcon, X, Trophy, Medal, QrCode, Copy, UserCheck,
  Eye, EyeOff, MousePointerClick, CalendarPlus, Calendar, MapPin, Mail, Maximize2, ClipboardList, Briefcase, LayoutGrid,
} from "lucide-react";
import {
  RadarChart, PolarGrid, PolarAngleAxis, Radar, ResponsiveContainer,
  LineChart, Line, XAxis, YAxis, CartesianGrid, ReferenceLine, Tooltip,
  BarChart, Bar, PieChart, Pie, Cell, Legend,
} from "recharts";

// Shared storage backed by a serverless API route (see /api/data.js), which
// itself is backed by Vercel KV. This replaces both Claude's artifact-only
// `window.storage` and any local-only localStorage approach, so every
// device that opens the deployed site reads and writes the same data.
const storage = {
  async get(key) {
    const res = await fetch(`/api/data?key=${encodeURIComponent(key)}`);
    if (!res.ok) throw new Error("not found");
    const data = await res.json();
    return { key, value: data.value };
  },
  async set(key, value) {
    const res = await fetch(`/api/data?key=${encodeURIComponent(key)}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ value }),
    });
    if (!res.ok) throw new Error("save failed");
    return { key, value };
  },
  async delete(key) {
    await fetch(`/api/data?key=${encodeURIComponent(key)}`, { method: "DELETE" });
    return { key, deleted: true };
  },
};

// ---------------------------------------------------------------------------
// Mock data — edit these objects to change what's shown across the app
// ---------------------------------------------------------------------------

// `self: true` marks the person viewing Dashboard / My Profile. Everyone else
// shows up only in the Admin participant list. This is now seed data for
// live, editable React state — see AppDataProvider below.
const participantsSeedRaw = [
  { employeeId: "EMP-4892", name: "Mokhtar Alkhlifa", email: "Mokhtar.Alkhlifa@acmecorp.com", cohort: "Cohort B", dept: "Strategy & Innovation", enrolled: "2024-09-01", count: 4, last: "2026-03-10", status: "Active", avgScore: 81, self: true, passwordSalt: "demoSalt123", passwordHash: "10a71203c4279d0fd55eacf9642862f1e491bf977af5b55653c9f42971c9f3a3" },
  { employeeId: "EMP-3341", name: "James Okafor", email: "j.okafor@globaltech.com", cohort: "Cohort A", dept: "Operations", enrolled: "2025-02-14", count: 3, last: "2026-02-28", status: "Active", avgScore: 88, passwordSalt: null, passwordHash: null },
  { employeeId: "EMP-5510", name: "Priya Nair", email: "p.nair@finserv.io", cohort: "Cohort C", dept: "Finance", enrolled: "2025-04-02", count: 2, last: "2026-01-15", status: "Active", avgScore: 92, passwordSalt: null, passwordHash: null },
  { employeeId: "EMP-2278", name: "Marcus Chen", email: "m.chen@techcorp.com", cohort: "Cohort B", dept: "Technology", enrolled: "2025-01-20", count: 3, last: "2026-03-05", status: "Active", avgScore: 76, passwordSalt: null, passwordHash: null },
  { employeeId: "EMP-6634", name: "Amara Diallo", email: "a.diallo@consulting.co", cohort: "Cohort A", dept: "Consulting", enrolled: "2024-11-08", count: 1, last: "2025-12-10", status: "Inactive", avgScore: 65, passwordSalt: null, passwordHash: null },
  { employeeId: "EMP-7712", name: "Lena Fischer", email: "l.fischer@eurocorp.de", cohort: "Cohort C", dept: "HR", enrolled: "2025-03-18", count: 2, last: "2026-02-20", status: "Active", avgScore: 94, passwordSalt: null, passwordHash: null },
  { employeeId: "EMP-8823", name: "David Park", email: "d.park@innovate.kr", cohort: "Cohort B", dept: "Product", enrolled: "2024-08-30", count: 4, last: "2026-03-08", status: "Active", avgScore: 83, passwordSalt: null, passwordHash: null },
  { employeeId: "EMP-9012", name: "Tariq Hassan", email: "t.hassan@brightline.io", cohort: "Cohort A", dept: "Marketing", enrolled: "2025-06-01", count: 2, last: "2026-01-30", status: "Active", avgScore: 71, passwordSalt: null, passwordHash: null },
  { employeeId: "EMP-1001", name: "Mukhtar Alkhlifa", email: "slark820@gmail.com", cohort: "Cohort A", dept: "Product", enrolled: new Date().toISOString().slice(0, 10), count: 0, last: new Date().toISOString().slice(0, 10), status: "Active", avgScore: 0, passwordSalt: null, passwordHash: null },
];
// Everyone in the built-in sample data belongs to the default project, so
// existing behavior is unchanged until an admin creates a second project.
const participantsSeed = participantsSeedRaw.map((p) => ({ projectId: "default", ...p }));

const assessments = [
  {
    id: "leadership",
    name: "Leadership Foundations",
    program: "Mercer Leadership Excellence Program 2025",
    submitted: "2025-11-14",
    graded: "2025-11-18",
    score: 84,
    status: "Pass",
    certificate: true,
    criteria: [
      { name: "Innovation", weight: 25, score: 88 },
      { name: "Critical Thinking", weight: 20, score: 80 },
      { name: "Clarity", weight: 20, score: 85 },
      { name: "Depth of Analysis", weight: 25, score: 82 },
      { name: "Grammar", weight: 10, score: 90 },
    ],
  },
  {
    id: "digital",
    name: "Digital Transformation Strategy",
    program: "Mercer Digital Accelerator Program 2025",
    submitted: "2025-12-02",
    graded: "2025-12-08",
    score: 91,
    status: "Pass",
    certificate: true,
    criteria: [
      { name: "Innovation", weight: 25, score: 95 },
      { name: "Critical Thinking", weight: 20, score: 89 },
      { name: "Clarity", weight: 20, score: 92 },
      { name: "Depth of Analysis", weight: 25, score: 90 },
      { name: "Grammar", weight: 10, score: 94 },
    ],
  },
  {
    id: "change",
    name: "Change Management Essentials",
    program: "Mercer HR Business Partner Track 2026",
    submitted: "2026-01-20",
    graded: "2026-01-24",
    score: 81,
    status: "Pass",
    certificate: true,
    criteria: [
      { name: "Innovation", weight: 25, score: 78 },
      { name: "Critical Thinking", weight: 20, score: 80 },
      { name: "Clarity", weight: 20, score: 85 },
      { name: "Depth of Analysis", weight: 25, score: 80 },
      { name: "Grammar", weight: 10, score: 88 },
    ],
    feedback: {
      strongest: { name: "Grammar", score: 88 },
      growth: { name: "Innovation", score: 78 },
      notes: [
        {
          name: "Innovation",
          score: 78,
          text: "The reflection showed solid creative thinking in the proposed change communication plan, drawing on established models and adapting them well to the specific organizational context described.",
        },
        {
          name: "Critical Thinking",
          score: 80,
          text: "Strong critical analysis with well-supported arguments. The participant identified the right challenges and followed the reasoning through to meaningful, actionable conclusions.",
        },
      ],
    },
  },
  {
    id: "workforce",
    name: "Strategic Workforce Planning",
    program: "Mercer HR Business Partner Track 2026",
    submitted: "2026-03-10",
    graded: "2026-03-10",
    score: null,
    status: "Grading",
    certificate: false,
    reflection:
      "In today's rapidly evolving business landscape, strategic workforce planning has become a cornerstone of organizational resilience. Through this program, I have developed a deeper understanding of how talent pipelines must be aligned with long-term business objectives rather than short-term operational needs. My key insight is that workforce planning is not merely a HR function — it is a strategic lever that, when used effectively, can differentiate organizations in competitive markets…",
    criteria: [
      { name: "Innovation", weight: 25, score: 82 },
      { name: "Critical Thinking", weight: 20, score: 76 },
      { name: "Clarity", weight: 20, score: 80 },
      { name: "Depth of Analysis", weight: 25, score: 74 },
      { name: "Grammar", weight: 10, score: 88 },
    ],
  },
];

const passingThreshold = 70;

const adminInfo = { email: "admin.grader@mercer.com", totalParticipants: 24, pendingReview: 3, certsIssued: 18 };

const surveysSeed = [
  { id: "q1-2026", name: "Program Feedback Survey — Q1 2026", responses: 186, rate: 78, uploaded: "2026-03-12", sourceType: "upload" },
  { id: "onboarding", name: "New Hire Onboarding Pulse", responses: 94, rate: 61, uploaded: "2026-02-20", sourceType: "upload" },
  { id: "leadership-nps", name: "Leadership Program NPS", responses: 142, rate: 85, uploaded: "2026-01-08", sourceType: "upload" },
  { id: "cbuae-pre-impact", name: "Pre Impact Survey — Leadership Program", responses: 21, rate: 100, uploaded: "2025-09-02", sourceType: "upload", fileName: "Pre_Impact_Survey_assessment-1000001154-raw-results_09-02-2025.xlsx" },
  { id: "cbuae-post-impact", name: "Post Impact Survey — Leadership Program", responses: 19, rate: 90, uploaded: "2025-12-10", sourceType: "upload", fileName: "Post_Impact_Survey_assessment-1000001154-raw-results_12-10-2025.xlsx" },
];

const surveyRatingData = [
  { question: "Content quality", avg: 4.4 },
  { question: "Facilitator", avg: 4.6 },
  { question: "Pace", avg: 3.9 },
  { question: "Relevance", avg: 4.2 },
  { question: "Materials", avg: 4.0 },
];

const surveyChoiceData = [
  { name: "Very likely", value: 96, color: "#0d9488" },
  { name: "Likely", value: 54, color: "#38bdf8" },
  { name: "Neutral", value: 24, color: "#fbbf24" },
  { name: "Unlikely", value: 12, color: "#f97316" },
];

const surveyKeywords = [
  { word: "engaging facilitators", weight: 5 },
  { word: "too fast-paced", weight: 4 },
  { word: "actionable takeaways", weight: 4 },
  { word: "wanted more case studies", weight: 3 },
  { word: "great group discussions", weight: 3 },
  { word: "scheduling conflicts", weight: 2 },
];

// Per-survey analysis data. Surveys not listed here fall back to the generic
// mock arrays above. Computed from the actual uploaded file's Likert-scale
// responses (19 completed respondents) — question text is preserved since
// it's the substance of the survey, but no individual respondent names or
// emails are included anywhere in this app, even though the source file
// contained them. Aggregate only.
const surveyAnalysisSeed = {
  "cbuae-pre-impact": {
    ratingLabel: "Rating by Question — Average Score (1–5) · top 12 of 29",
    ratingData: [
      { question: "I am capable of fostering a sense of urgency in my team/…", fullQuestion: "I am capable of fostering a sense of urgency in my team/department to reach goals and meet deadlines", avg: 5.0 },
      { question: "I am capable of pushing my team/department to achieve de…", fullQuestion: "I am capable of pushing my team/department to achieve desirable results", avg: 4.84 },
      { question: "I am capable of acting in the best interest of others an…", fullQuestion: "I am capable of acting in the best interest of others and CBUAE even when it conflicts with my self-interest", avg: 4.84 },
      { question: "I feel confident in my ability to motivate and inspire m…", fullQuestion: "I feel confident in my ability to motivate and inspire my team to achieve their goals even during challenging times", avg: 4.84 },
      { question: "I am aware of the strengths, weaknesses, threats and opp…", fullQuestion: "I am aware of the strengths, weaknesses, threats and opportunities of my team/department", avg: 4.84 },
      { question: "I believe I have the skills and leadership qualities to …", fullQuestion: "I believe I have the skills and leadership qualities to progress in my career", avg: 4.8 },
      { question: "I am honest and transparent, my words and actions are al…", fullQuestion: "I am honest and transparent, my words and actions are aligned", avg: 4.8 },
      { question: "I am capable of creating a psychologically safe environm…", fullQuestion: "I am capable of creating a psychologically safe environment to enable my team/department to voice their concerns without retribution in mind", avg: 4.8 },
      { question: "I feel confident engaging in strategic decision making", fullQuestion: "I feel confident engaging in strategic decision making", avg: 4.68 },
      { question: "I am aware of how to leverage my strengths in the workpl…", fullQuestion: "I am aware of how to leverage my strengths in the workplace", avg: 4.68 },
      { question: "I am capable of displaying stability and composure despi…", fullQuestion: "I am capable of displaying stability and composure despite stressful and crisis situations", avg: 4.68 },
      { question: "I am capable of leading change from within while keeping…", fullQuestion: "I am capable of leading change from within while keeping all stakeholders aligned", avg: 4.68 },
    ],
    choiceTitle: "Overall Sentiment Across All Statements",
    choiceSubtitle: "609 individual responses across 29 statements",
    choiceData: [
      { name: "Strongly Agree", value: 264, color: "#0d9488" },
      { name: "Agree", value: 264, color: "#38bdf8" },
      { name: "Neutral", value: 81, color: "#fbbf24" },
    ],
    distributionData: [
      { question: "I am capable of fostering a sense of urgency…", fullQuestion: "I am capable of fostering a sense of urgency in my team/department to reach goals and meet deadlines", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 0, "Agree": 50.0, "Strongly Agree": 50.0 },
      { question: "I am capable of pushing my team/department t…", fullQuestion: "I am capable of pushing my team/department to achieve desirable results", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 4, "Agree": 48.0, "Strongly Agree": 48.0 },
      { question: "I am capable of acting in the best interest …", fullQuestion: "I am capable of acting in the best interest of others and CBUAE even when it conflicts with my self-interest", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 4, "Agree": 48.0, "Strongly Agree": 48.0 },
      { question: "I feel confident in my ability to motivate a…", fullQuestion: "I feel confident in my ability to motivate and inspire my team to achieve their goals even during challenging times", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 4, "Agree": 48.0, "Strongly Agree": 48.0 },
      { question: "I am aware of the strengths, weaknesses, thr…", fullQuestion: "I am aware of the strengths, weaknesses, threats and opportunities of my team/department", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 4, "Agree": 48.0, "Strongly Agree": 48.0 },
      { question: "I believe I have the skills and leadership q…", fullQuestion: "I believe I have the skills and leadership qualities to progress in my career", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 5, "Agree": 47.5, "Strongly Agree": 47.5 },
      { question: "I am honest and transparent, my words and ac…", fullQuestion: "I am honest and transparent, my words and actions are aligned", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 5, "Agree": 47.5, "Strongly Agree": 47.5 },
      { question: "I am capable of creating a psychologically s…", fullQuestion: "I am capable of creating a psychologically safe environment to enable my team/department to voice their concerns without retribution in mind", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 5, "Agree": 47.5, "Strongly Agree": 47.5 },
      { question: "I feel confident engaging in strategic decis…", fullQuestion: "I feel confident engaging in strategic decision making", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 8, "Agree": 46.0, "Strongly Agree": 46.0 },
      { question: "I am aware of how to leverage my strengths i…", fullQuestion: "I am aware of how to leverage my strengths in the workplace", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 8, "Agree": 46.0, "Strongly Agree": 46.0 },
      { question: "I am capable of displaying stability and com…", fullQuestion: "I am capable of displaying stability and composure despite stressful and crisis situations", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 8, "Agree": 46.0, "Strongly Agree": 46.0 },
      { question: "I am capable of leading change from within w…", fullQuestion: "I am capable of leading change from within while keeping all stakeholders aligned", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 8, "Agree": 46.0, "Strongly Agree": 46.0 },
      { question: "I am able to think and respond from a strate…", fullQuestion: "I am able to think and respond from a strategic, bigger picture perspective", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 8, "Agree": 46.0, "Strongly Agree": 46.0 },
      { question: "I feel confident setting strategy for my tea…", fullQuestion: "I feel confident setting strategy for my team/department", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 12, "Agree": 44.0, "Strongly Agree": 44.0 },
      { question: "I am aware of how to build my personal resil…", fullQuestion: "I am aware of how to build my personal resilience", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 12, "Agree": 44.0, "Strongly Agree": 44.0 },
      { question: "I am comfortable in networking and communica…", fullQuestion: "I am comfortable in networking and communicating with more senior stakeholders", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 12, "Agree": 44.0, "Strongly Agree": 44.0 },
      { question: "I feel confident making long term decisions …", fullQuestion: "I feel confident making long term decisions for CBUAE", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 12, "Agree": 44.0, "Strongly Agree": 44.0 },
      { question: "I am capable of optimizing decisions across …", fullQuestion: "I am capable of optimizing decisions across the varied needs of numerous stakeholders", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 12, "Agree": 44.0, "Strongly Agree": 44.0 },
      { question: "I know what I need to do to develop myself i…", fullQuestion: "I know what I need to do to develop myself in the future", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 16, "Agree": 42.0, "Strongly Agree": 42.0 },
      { question: "I feel confident in my ability to communicat…", fullQuestion: "I feel confident in my ability to communicate with others and express my views", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 16, "Agree": 42.0, "Strongly Agree": 42.0 },
      { question: "I can effectively deal with conflict within …", fullQuestion: "I can effectively deal with conflict within teams", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 16, "Agree": 42.0, "Strongly Agree": 42.0 },
      { question: "I can effectively read and understand other …", fullQuestion: "I can effectively read and understand other people's emotions and feelings", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 16, "Agree": 42.0, "Strongly Agree": 42.0 },
      { question: "I believe I am capable of challenging others…", fullQuestion: "I believe I am capable of challenging others effectively", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 16, "Agree": 42.0, "Strongly Agree": 42.0 },
      { question: "I am capable of resolving conflict, receivin…", fullQuestion: "I am capable of resolving conflict, receiving resources, and handling challenges through my internal network at CBUAE", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 20, "Agree": 40.0, "Strongly Agree": 40.0 },
      { question: "I believe I am comfortable with navigating u…", fullQuestion: "I believe I am comfortable with navigating uncertainty and adapting strategies to effectively lead my team/department through change", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 20, "Agree": 40.0, "Strongly Agree": 40.0 },
      { question: "I am capable of demonstrating foresight abou…", fullQuestion: "I am capable of demonstrating foresight about how stakeholder priorities will change and evolve", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 20, "Agree": 40.0, "Strongly Agree": 40.0 },
      { question: "I feel comfortable disclosing about my genui…", fullQuestion: "I feel comfortable disclosing about my genuine self at work and my realities, challenges and successes", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 24, "Agree": 38.0, "Strongly Agree": 38.0 },
      { question: "I have the required guidance and support I n…", fullQuestion: "I have the required guidance and support I need to develop in my career", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 44, "Agree": 28.0, "Strongly Agree": 28.0 },
      { question: "I feel that I am generally comfortable with …", fullQuestion: "I feel that I am generally comfortable with public speaking, i.e. being able to present and speak in front of others", "Strongly Disagree": 0, "Disagree": 0, "Neutral": 44, "Agree": 28.0, "Strongly Agree": 28.0 },
    ],
    fullQuestionList: [
      { statement: "I am capable of fostering a sense of urgency in my team/department to reach goals and meet deadlines", avg: 5.0, n: 21 },
      { statement: "I am capable of pushing my team/department to achieve desirable results", avg: 4.84, n: 21 },
      { statement: "I am capable of acting in the best interest of others and CBUAE even when it conflicts with my self-interest", avg: 4.84, n: 21 },
      { statement: "I feel confident in my ability to motivate and inspire my team to achieve their goals even during challenging times", avg: 4.84, n: 21 },
      { statement: "I am aware of the strengths, weaknesses, threats and opportunities of my team/department", avg: 4.84, n: 21 },
      { statement: "I believe I have the skills and leadership qualities to progress in my career", avg: 4.8, n: 21 },
      { statement: "I am honest and transparent, my words and actions are aligned", avg: 4.8, n: 21 },
      { statement: "I am capable of creating a psychologically safe environment to enable my team/department to voice their concerns without retribution in mind", avg: 4.8, n: 21 },
      { statement: "I feel confident engaging in strategic decision making", avg: 4.68, n: 21 },
      { statement: "I am aware of how to leverage my strengths in the workplace", avg: 4.68, n: 21 },
      { statement: "I am capable of displaying stability and composure despite stressful and crisis situations", avg: 4.68, n: 21 },
      { statement: "I am capable of leading change from within while keeping all stakeholders aligned", avg: 4.68, n: 21 },
      { statement: "I am able to think and respond from a strategic, bigger picture perspective", avg: 4.68, n: 21 },
      { statement: "I feel confident setting strategy for my team/department", avg: 4.52, n: 21 },
      { statement: "I am aware of how to build my personal resilience", avg: 4.52, n: 21 },
      { statement: "I am comfortable in networking and communicating with more senior stakeholders", avg: 4.52, n: 21 },
      { statement: "I feel confident making long term decisions for CBUAE", avg: 4.52, n: 21 },
      { statement: "I am capable of optimizing decisions across the varied needs of numerous stakeholders", avg: 4.52, n: 21 },
      { statement: "I know what I need to do to develop myself in the future", avg: 4.36, n: 21 },
      { statement: "I feel confident in my ability to communicate with others and express my views", avg: 4.36, n: 21 },
      { statement: "I can effectively deal with conflict within teams", avg: 4.36, n: 21 },
      { statement: "I can effectively read and understand other people's emotions and feelings", avg: 4.36, n: 21 },
      { statement: "I believe I am capable of challenging others effectively", avg: 4.36, n: 21 },
      { statement: "I am capable of resolving conflict, receiving resources, and handling challenges through my internal network at CBUAE", avg: 4.2, n: 21 },
      { statement: "I believe I am comfortable with navigating uncertainty and adapting strategies to effectively lead my team/department through change", avg: 4.2, n: 21 },
      { statement: "I am capable of demonstrating foresight about how stakeholder priorities will change and evolve", avg: 4.2, n: 21 },
      { statement: "I feel comfortable disclosing about my genuine self at work and my realities, challenges and successes", avg: 4.04, n: 21 },
      { statement: "I have the required guidance and support I need to develop in my career", avg: 3.24, n: 21 },
      { statement: "I feel that I am generally comfortable with public speaking, i.e. being able to present and speak in front of others", avg: 3.24, n: 21 },
    ],
    crossTabTitle: "Highest & Lowest Rated Statements",
    crossTabRows: [
      { statement: "I am capable of fostering a sense of urgency in my team/department to reach goals and meet deadlines", avg: 5.0, band: "top" },
      { statement: "I am capable of pushing my team/department to achieve desirable results", avg: 4.84, band: "top" },
      { statement: "I am capable of acting in the best interest of others and CBUAE even when it conflicts with my self-interest", avg: 4.84, band: "top" },
      { statement: "I feel confident in my ability to motivate and inspire my team to achieve their goals even during challenging times", avg: 4.84, band: "top" },
      { statement: "I am aware of the strengths, weaknesses, threats and opportunities of my team/department", avg: 4.84, band: "top" },
      { statement: "I have the required guidance and support I need to develop in my career", avg: 3.24, band: "bottom" },
      { statement: "I feel that I am generally comfortable with public speaking, i.e. being able to present and speak in front of others", avg: 3.24, band: "bottom" },
      { statement: "I feel comfortable disclosing about my genuine self at work and my realities, challenges and successes", avg: 4.04, band: "bottom" },
      { statement: "I am capable of resolving conflict, receiving resources, and handling challenges through my internal network at CBUAE", avg: 4.2, band: "bottom" },
      { statement: "I believe I am comfortable with navigating uncertainty and adapting strategies to effectively lead my team/department through change", avg: 4.2, band: "bottom" },
    ],
    keywords: null,
  },
  "cbuae-post-impact": {
    ratingLabel: "Rating by Theme — Average Score (1–5)",
    ratingData: [
      { question: "Strategic Thinking", avg: 4.58 },
      { question: "Self-Dev & Resilience", avg: 4.68 },
      { question: "Communication", avg: 4.51 },
      { question: "Career Growth", avg: 4.61 },
      { question: "Courage & Integrity", avg: 4.58 },
      { question: "Change Leadership", avg: 4.54 },
    ],
    choiceTitle: "Overall Sentiment Across All Statements",
    choiceSubtitle: "551 individual responses across 29 statements",
    choiceData: [
      { name: "Strongly Agree", value: 353, color: "#0d9488" },
      { name: "Agree", value: 165, color: "#38bdf8" },
      { name: "Neutral", value: 30, color: "#fbbf24" },
      { name: "Disagree", value: 2, color: "#f97316" },
      { name: "Strongly Disagree", value: 1, color: "#dc2626" },
    ],
    distributionData: [
      { question: "I know what I need to do to develop my…", fullQuestion: "I know what I need to do to develop myself in the future", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 0.0, "Agree": 15.8, "Strongly Agree": 84.2 },
      { question: "I feel confident setting strategy for …", fullQuestion: "I feel confident setting strategy for my team/department", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 15.8, "Strongly Agree": 78.9 },
      { question: "I feel confident engaging in strategic…", fullQuestion: "I feel confident engaging in strategic decision making", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 15.8, "Strongly Agree": 78.9 },
      { question: "I am capable of pushing my team/depart…", fullQuestion: "I am capable of pushing my team/department to achieve desirable results", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 15.8, "Strongly Agree": 78.9 },
      { question: "I am aware of how to leverage my stren…", fullQuestion: "I am aware of how to leverage my strengths in the workplace", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 0.0, "Agree": 31.6, "Strongly Agree": 68.4 },
      { question: "I believe I have the skills and leader…", fullQuestion: "I believe I have the skills and leadership qualities to progress in my career", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 21.1, "Strongly Agree": 73.7 },
      { question: "I am capable of acting in the best int…", fullQuestion: "I am capable of acting in the best interest of others and CBUAE even when it conflicts with my self-interest", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 21.1, "Strongly Agree": 73.7 },
      { question: "I am aware of how to build my personal…", fullQuestion: "I am aware of how to build my personal resilience", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 26.3, "Strongly Agree": 68.4 },
      { question: "I feel comfortable disclosing about my…", fullQuestion: "I feel comfortable disclosing about my genuine self at work and my realities, challenges and successes", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 10.5, "Agree": 15.8, "Strongly Agree": 73.7 },
      { question: "I am honest and transparent, my words …", fullQuestion: "I am honest and transparent, my words and actions are aligned", "Strongly Disagree": 0.0, "Disagree": 5.3, "Neutral": 0.0, "Agree": 21.1, "Strongly Agree": 73.7 },
      { question: "I feel confident in my ability to moti…", fullQuestion: "I feel confident in my ability to motivate and inspire my team to achieve their goals even during challenging times", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 26.3, "Strongly Agree": 68.4 },
      { question: "I am capable of displaying stability a…", fullQuestion: "I am capable of displaying stability and composure despite stressful and crisis situations", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 31.6, "Strongly Agree": 63.2 },
      { question: "I feel confident in my ability to comm…", fullQuestion: "I feel confident in my ability to communicate with others and express my views", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 31.6, "Strongly Agree": 63.2 },
      { question: "I can effectively deal with conflict w…", fullQuestion: "I can effectively deal with conflict within teams", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 31.6, "Strongly Agree": 63.2 },
      { question: "I am comfortable in networking and com…", fullQuestion: "I am comfortable in networking and communicating with more senior stakeholders", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 10.5, "Agree": 21.1, "Strongly Agree": 68.4 },
      { question: "I am capable of resolving conflict, re…", fullQuestion: "I am capable of resolving conflict, receiving resources, and handling challenges through my internal network at CBUAE", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 31.6, "Strongly Agree": 63.2 },
      { question: "I have the required guidance and suppo…", fullQuestion: "I have the required guidance and support I need to develop in my career", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 0.0, "Agree": 47.4, "Strongly Agree": 52.6 },
      { question: "I can effectively read and understand …", fullQuestion: "I can effectively read and understand other people's emotions and feelings", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 0.0, "Agree": 47.4, "Strongly Agree": 52.6 },
      { question: "I am aware of the strengths, weaknesse…", fullQuestion: "I am aware of the strengths, weaknesses, threats and opportunities of my team/department", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 36.8, "Strongly Agree": 57.9 },
      { question: "I am capable of creating a psychologic…", fullQuestion: "I am capable of creating a psychologically safe environment to enable my team/department to voice their concerns without retribution in mind", "Strongly Disagree": 0.0, "Disagree": 5.3, "Neutral": 0.0, "Agree": 31.6, "Strongly Agree": 63.2 },
      { question: "I feel confident making long term deci…", fullQuestion: "I feel confident making long term decisions for CBUAE", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 10.5, "Agree": 31.6, "Strongly Agree": 57.9 },
      { question: "I believe I am capable of challenging …", fullQuestion: "I believe I am capable of challenging others effectively", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 10.5, "Agree": 31.6, "Strongly Agree": 57.9 },
      { question: "I am capable of leading change from wi…", fullQuestion: "I am capable of leading change from within while keeping all stakeholders aligned", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 10.5, "Agree": 31.6, "Strongly Agree": 57.9 },
      { question: "I believe I am comfortable with naviga…", fullQuestion: "I believe I am comfortable with navigating uncertainty and adapting strategies to effectively lead my team/department through change", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 42.1, "Strongly Agree": 52.6 },
      { question: "I am capable of demonstrating foresigh…", fullQuestion: "I am capable of demonstrating foresight about how stakeholder priorities will change and evolve", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 10.5, "Agree": 31.6, "Strongly Agree": 57.9 },
      { question: "I feel that I am generally comfortable…", fullQuestion: "I feel that I am generally comfortable with public speaking, i.e. being able to present and speak in front of others", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 47.4, "Strongly Agree": 47.4 },
      { question: "I am capable of fostering a sense of u…", fullQuestion: "I am capable of fostering a sense of urgency in my team/department to reach goals and meet deadlines", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 5.3, "Agree": 47.4, "Strongly Agree": 47.4 },
      { question: "I am capable of optimizing decisions a…", fullQuestion: "I am capable of optimizing decisions across the varied needs of numerous stakeholders", "Strongly Disagree": 0.0, "Disagree": 0.0, "Neutral": 10.5, "Agree": 36.8, "Strongly Agree": 52.6 },
      { question: "I am able to think and respond from a …", fullQuestion: "I am able to think and respond from a strategic, bigger picture perspective", "Strongly Disagree": 5.3, "Disagree": 0.0, "Neutral": 5.3, "Agree": 31.6, "Strongly Agree": 57.9 },
    ],
    fullQuestionList: [
      { statement: "I know what I need to do to develop myself in the future", avg: 4.84, n: 19 },
      { statement: "I feel confident setting strategy for my team/department", avg: 4.74, n: 19 },
      { statement: "I feel confident engaging in strategic decision making", avg: 4.74, n: 19 },
      { statement: "I am capable of pushing my team/department to achieve desirable results", avg: 4.74, n: 19 },
      { statement: "I am aware of how to leverage my strengths in the workplace", avg: 4.68, n: 19 },
      { statement: "I believe I have the skills and leadership qualities to progress in my career", avg: 4.68, n: 19 },
      { statement: "I am capable of acting in the best interest of others and CBUAE even when it conflicts with my self-interest", avg: 4.68, n: 19 },
      { statement: "I am aware of how to build my personal resilience", avg: 4.63, n: 19 },
      { statement: "I feel comfortable disclosing about my genuine self at work and my realities, challenges and successes", avg: 4.63, n: 19 },
      { statement: "I am honest and transparent, my words and actions are aligned", avg: 4.63, n: 19 },
      { statement: "I feel confident in my ability to motivate and inspire my team to achieve their goals even during challenging times", avg: 4.63, n: 19 },
      { statement: "I am capable of displaying stability and composure despite stressful and crisis situations", avg: 4.58, n: 19 },
      { statement: "I feel confident in my ability to communicate with others and express my views", avg: 4.58, n: 19 },
      { statement: "I can effectively deal with conflict within teams", avg: 4.58, n: 19 },
      { statement: "I am comfortable in networking and communicating with more senior stakeholders", avg: 4.58, n: 19 },
      { statement: "I am capable of resolving conflict, receiving resources, and handling challenges through my internal network at CBUAE", avg: 4.58, n: 19 },
      { statement: "I have the required guidance and support I need to develop in my career", avg: 4.53, n: 19 },
      { statement: "I can effectively read and understand other people's emotions and feelings", avg: 4.53, n: 19 },
      { statement: "I am aware of the strengths, weaknesses, threats and opportunities of my team/department", avg: 4.53, n: 19 },
      { statement: "I am capable of creating a psychologically safe environment to enable my team/department to voice their concerns without retribution in mind", avg: 4.53, n: 19 },
      { statement: "I feel confident making long term decisions for CBUAE", avg: 4.47, n: 19 },
      { statement: "I believe I am capable of challenging others effectively", avg: 4.47, n: 19 },
      { statement: "I am capable of leading change from within while keeping all stakeholders aligned", avg: 4.47, n: 19 },
      { statement: "I believe I am comfortable with navigating uncertainty and adapting strategies to effectively lead my team/department through change", avg: 4.47, n: 19 },
      { statement: "I am capable of demonstrating foresight about how stakeholder priorities will change and evolve", avg: 4.47, n: 19 },
      { statement: "I feel that I am generally comfortable with public speaking, i.e. being able to present and speak in front of others", avg: 4.42, n: 19 },
      { statement: "I am capable of fostering a sense of urgency in my team/department to reach goals and meet deadlines", avg: 4.42, n: 19 },
      { statement: "I am capable of optimizing decisions across the varied needs of numerous stakeholders", avg: 4.42, n: 19 },
      { statement: "I am able to think and respond from a strategic, bigger picture perspective", avg: 4.37, n: 19 },
    ],
    crossTabTitle: "Highest & Lowest Rated Statements",
    crossTabRows: [
      { statement: "I know what I need to do to develop myself in the future", avg: 4.84, band: "top" },
      { statement: "I feel confident setting strategy for my team/department", avg: 4.74, band: "top" },
      { statement: "I feel confident engaging in strategic decision making", avg: 4.74, band: "top" },
      { statement: "I am capable of pushing my team/department to achieve desirable results", avg: 4.74, band: "top" },
      { statement: "I am aware of how to leverage my strengths in the workplace", avg: 4.68, band: "top" },
      { statement: "I am able to think and respond from a strategic, bigger picture perspective", avg: 4.37, band: "bottom" },
      { statement: "I am capable of optimizing decisions across the varied needs of numerous stakeholders", avg: 4.42, band: "bottom" },
      { statement: "I am capable of fostering a sense of urgency in my team/department to reach goals and meet deadlines", avg: 4.42, band: "bottom" },
      { statement: "I feel comfortable with public speaking, presenting in front of others", avg: 4.42, band: "bottom" },
      { statement: "I am capable of demonstrating foresight about how stakeholder priorities will change", avg: 4.47, band: "bottom" },
    ],
    keywords: null, // this survey has no open-text questions to extract themes from
  },
};

// ---------------------------------------------------------------------------
// Real survey file analysis — runs entirely client-side on whatever the user
// drops in. Looks for a response sheet with participant name/email columns
// followed by question columns (matching common survey-export shapes like
// "Subject Name / Subject Email / Status / Q1 / Q2 / ..." with an optional
// full-question-text row beneath the header). Likert-style answers are
// mapped to a 1-5 scale; columns that are mostly longer free text are
// treated as open-ended and get lightweight keyword extraction instead.
// Never reads or stores individual respondent names/emails anywhere in the
// output — only aggregate statistics.
// ---------------------------------------------------------------------------

const LIKERT_MAP = {
  "strongly disagree": 1, "disagree": 2, "neutral": 3, "agree": 4, "strongly agree": 5,
  "very unlikely": 1, "unlikely": 2, "somewhat likely": 3, "likely": 4, "very likely": 5,
  "poor": 1, "fair": 2, "good": 3, "very good": 4, "excellent": 5,
  "never": 1, "rarely": 2, "sometimes": 3, "often": 4, "always": 5,
};
const LIKERT_LABELS_5 = ["Strongly Disagree", "Disagree", "Neutral", "Agree", "Strongly Agree"];
const LIKERT_COLORS_5 = ["#dc2626", "#f97316", "#fbbf24", "#38bdf8", "#0d9488"];
const SURVEY_STOPWORDS = new Set([
  "the", "a", "an", "and", "or", "but", "is", "are", "was", "were", "to", "of", "in", "on", "for",
  "with", "that", "this", "it", "as", "at", "by", "be", "i", "my", "me", "we", "our", "you", "your",
  "have", "has", "had", "not", "no", "so", "if", "than", "then", "also", "very", "more", "most",
  "much", "just", "can", "could", "would", "should", "will", "from", "about", "into", "which",
  "what", "who", "when", "where", "there", "their", "them", "these", "those", "been", "being",
]);

function likertValue(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") return raw >= 1 && raw <= 5 ? raw : null;
  const s = String(raw).trim().toLowerCase();
  return s in LIKERT_MAP ? LIKERT_MAP[s] : null;
}

function analyzeSurveyWorkbook(workbook) {
  let best = null;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, blankrows: false });
    if (!rows || rows.length < 2) continue;

    let headerRowIdx = -1, nameCol = -1, emailCol = -1, statusCol = -1;
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r] || [];
      let nc = -1, ec = -1, sc = -1;
      row.forEach((cell, ci) => {
        if (cell == null) return;
        const s = String(cell).trim().toLowerCase();
        if (nc === -1 && /(subject|participant|respondent)?\s*name/.test(s)) nc = ci;
        if (ec === -1 && /e-?mail/.test(s)) ec = ci;
        if (sc === -1 && /status/.test(s)) sc = ci;
      });
      if (nc !== -1 && ec !== -1) { headerRowIdx = r; nameCol = nc; emailCol = ec; statusCol = sc; break; }
    }
    if (headerRowIdx === -1) continue;
    if (!best || rows.length > best.rows.length) {
      best = { sheetName, rows, headerRowIdx, nameCol, emailCol, statusCol };
    }
  }

  if (!best) {
    throw new Error('Couldn\'t find a response sheet with participant name and email columns (looked for headers like "Subject Name" / "Subject Email"). Try a different file, or add this survey without a file.');
  }

  const { rows, headerRowIdx, nameCol, emailCol, statusCol } = best;
  const headerRow = rows[headerRowIdx] || [];
  const nextRow = rows[headerRowIdx + 1] || [];

  let longTextHits = 0, totalQuestionCols = 0;
  headerRow.forEach((cell, ci) => {
    if (ci === nameCol || ci === emailCol || ci === statusCol) return;
    if (cell == null || String(cell).trim() === "") return;
    totalQuestionCols++;
    const below = nextRow[ci];
    if (below != null && String(below).length > String(cell).length + 8) longTextHits++;
  });
  const hasTextRow = totalQuestionCols > 0 && longTextHits / totalQuestionCols > 0.5;
  const dataStartIdx = (hasTextRow ? headerRowIdx + 1 : headerRowIdx) + 1;

  const questionCols = [];
  headerRow.forEach((cell, ci) => {
    if (ci === nameCol || ci === emailCol || ci === statusCol) return;
    if (cell == null || String(cell).trim() === "") return;
    const raw = hasTextRow && nextRow[ci] ? String(nextRow[ci]) : String(cell);
    const label = raw.replace(/&nbsp;/gi, " ").replace(/\s+/g, " ").trim();
    questionCols.push({ index: ci, label });
  });
  if (questionCols.length === 0) {
    throw new Error("Found participant columns but no question columns after them to analyze.");
  }

  const dataRows = rows.slice(dataStartIdx).filter((r) => r && r[nameCol] != null && String(r[nameCol]).trim() !== "");
  if (dataRows.length === 0) {
    throw new Error("No respondent rows found under the detected header row.");
  }

  const completedRows = statusCol !== -1
    ? dataRows.filter((r) => {
        const s = r[statusCol] == null ? "" : String(r[statusCol]).trim().toLowerCase();
        return s === "" || s === "completed" || s === "complete" || s === "done";
      })
    : dataRows;
  const useRows = completedRows.length > 0 ? completedRows : dataRows;

  const likertCols = [], textCols = [];
  questionCols.forEach((q) => {
    let mappable = 0, present = 0;
    useRows.forEach((r) => {
      const v = r[q.index];
      if (v == null || String(v).trim() === "") return;
      present++;
      if (likertValue(v) != null) mappable++;
    });
    if (present === 0) return;
    if (mappable / present >= 0.7) likertCols.push(q);
    else textCols.push(q);
  });
  if (likertCols.length === 0 && textCols.length === 0) {
    throw new Error("Couldn't detect any ratable or open-text answers in this file's question columns.");
  }

  const qStats = likertCols
    .map((q) => {
      const vals = useRows.map((r) => likertValue(r[q.index])).filter((v) => v != null);
      return vals.length ? { ...q, avg: vals.reduce((s, v) => s + v, 0) / vals.length, vals } : null;
    })
    .filter(Boolean)
    .sort((a, b) => b.avg - a.avg);

  const ratingData = qStats.slice(0, 12).map((q) => ({
    question: q.label.length > 58 ? q.label.slice(0, 56) + "…" : q.label,
    fullQuestion: q.label,
    avg: Math.round(q.avg * 100) / 100,
  }));

  // Per-question distribution — each question's own answers as % of that
  // question's respondents, so questions can be compared and read
  // independently rather than pooled into one blended number.
  const distributionData = qStats.map((q) => {
    const counts = [0, 0, 0, 0, 0];
    q.vals.forEach((v) => { counts[Math.round(v) - 1]++; });
    const n = q.vals.length;
    const row = { question: q.label.length > 46 ? q.label.slice(0, 44) + "…" : q.label, fullQuestion: q.label };
    LIKERT_LABELS_5.forEach((label, i) => { row[label] = n ? Math.round((counts[i] / n) * 1000) / 10 : 0; });
    return row;
  });

  const fullQuestionList = qStats.map((q) => ({
    statement: q.label,
    avg: Math.round(q.avg * 100) / 100,
    n: q.vals.length,
  }));

  const bucketCounts = [0, 0, 0, 0, 0];
  qStats.forEach((q) => q.vals.forEach((v) => { bucketCounts[Math.round(v) - 1]++; }));
  const choiceData = LIKERT_LABELS_5
    .map((label, i) => ({ name: label, value: bucketCounts[i], color: LIKERT_COLORS_5[i] }))
    .filter((d) => d.value > 0);

  const top = qStats.slice(0, 5);
  const bottom = qStats.slice(-5).reverse();
  const crossTabRows = [
    ...top.map((q) => ({ statement: q.label, avg: q.avg, band: "top" })),
    ...bottom.filter((q) => !top.includes(q)).map((q) => ({ statement: q.label, avg: q.avg, band: "bottom" })),
  ];

  let keywords = null;
  if (textCols.length > 0) {
    const freq = new Map();
    useRows.forEach((r) => {
      textCols.forEach((q) => {
        const v = r[q.index];
        if (v == null) return;
        String(v).toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/)
          .filter((w) => w.length > 4 && !SURVEY_STOPWORDS.has(w))
          .forEach((w) => freq.set(w, (freq.get(w) || 0) + 1));
      });
    });
    const sorted = [...freq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8);
    if (sorted.length > 0) {
      const maxCount = sorted[0][1];
      keywords = sorted.map(([word, count]) => ({ word, weight: Math.max(1, Math.round((count / maxCount) * 5)) }));
    }
  }

  const totalRatings = bucketCounts.reduce((s, v) => s + v, 0);

  return {
    responses: useRows.length,
    rate: statusCol !== -1 && dataRows.length > 0 ? Math.round((useRows.length / dataRows.length) * 100) : 100,
    ratingLabel: `Rating by Question — Average Score (1–5)${qStats.length > 12 ? ` · top 12 of ${qStats.length}` : ""}`,
    ratingData,
    choiceTitle: "Overall Sentiment Across All Ratings",
    choiceSubtitle: `${totalRatings} individual responses across ${likertCols.length} question${likertCols.length === 1 ? "" : "s"}`,
    choiceData,
    distributionData,
    fullQuestionList,
    crossTabTitle: "Highest & Lowest Rated Statements",
    crossTabRows,
    keywords,
    questionCount: questionCols.length,
  };
}

// ---------------------------------------------------------------------------
// Analysis for admin-built surveys (real-time collected responses), sharing
// the exact same output shape as analyzeSurveyWorkbook() above so it plugs
// into all the same chart components with zero extra rendering logic.
// ---------------------------------------------------------------------------

function computeAnalysisFromResponses(questions, responseRows) {
  const ratingQs = (questions || []).filter((q) => q.type === "rating");
  const textQs = (questions || []).filter((q) => q.type === "text");

  const qStats = ratingQs
    .map((q) => {
      const vals = responseRows.map((r) => r.answers && r.answers[q.id]).filter((v) => typeof v === "number" && v >= 1 && v <= 5);
      return vals.length ? { label: q.text, avg: vals.reduce((s, v) => s + v, 0) / vals.length, vals } : null;
    })
    .filter(Boolean)
    .sort((a, b) => b.avg - a.avg);

  const ratingData = qStats.map((q) => ({
    question: q.label.length > 58 ? q.label.slice(0, 56) + "…" : q.label,
    fullQuestion: q.label,
    avg: Math.round(q.avg * 100) / 100,
  }));

  const distributionData = qStats.map((q) => {
    const counts = [0, 0, 0, 0, 0];
    q.vals.forEach((v) => { counts[Math.round(v) - 1]++; });
    const n = q.vals.length;
    const row = { question: q.label.length > 46 ? q.label.slice(0, 44) + "…" : q.label, fullQuestion: q.label };
    LIKERT_LABELS_5.forEach((label, i) => { row[label] = n ? Math.round((counts[i] / n) * 1000) / 10 : 0; });
    return row;
  });

  const bucketCounts = [0, 0, 0, 0, 0];
  qStats.forEach((q) => q.vals.forEach((v) => { bucketCounts[Math.round(v) - 1]++; }));
  const choiceData = LIKERT_LABELS_5
    .map((label, i) => ({ name: label, value: bucketCounts[i], color: LIKERT_COLORS_5[i] }))
    .filter((d) => d.value > 0);

  const fullQuestionList = qStats.map((q) => ({ statement: q.label, avg: Math.round(q.avg * 100) / 100, n: q.vals.length }));

  const top = qStats.slice(0, 5);
  const bottom = qStats.slice(-5).reverse();
  const crossTabRows = [
    ...top.map((q) => ({ statement: q.label, avg: q.avg, band: "top" })),
    ...bottom.filter((q) => !top.includes(q)).map((q) => ({ statement: q.label, avg: q.avg, band: "bottom" })),
  ];

  let keywords = null;
  if (textQs.length > 0) {
    const freq = new Map();
    responseRows.forEach((r) => {
      textQs.forEach((q) => {
        const v = r.answers && r.answers[q.id];
        if (!v) return;
        String(v).toLowerCase().replace(/[^a-z0-9\s]/g, " ").split(/\s+/)
          .filter((w) => w.length > 4 && !SURVEY_STOPWORDS.has(w))
          .forEach((w) => freq.set(w, (freq.get(w) || 0) + 1));
      });
    });
    const sorted = [...freq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8);
    if (sorted.length > 0) {
      const maxCount = sorted[0][1];
      keywords = sorted.map(([word, count]) => ({ word, weight: Math.max(1, Math.round((count / maxCount) * 5)) }));
    }
  }

  const totalRatings = bucketCounts.reduce((s, v) => s + v, 0);

  return {
    responses: responseRows.length,
    rate: 100,
    ratingLabel: "Rating by Question — Average Score (1–5)",
    ratingData,
    choiceTitle: "Overall Sentiment Across All Ratings",
    choiceSubtitle: `${totalRatings} individual responses across ${ratingQs.length} question${ratingQs.length === 1 ? "" : "s"}`,
    choiceData,
    distributionData,
    fullQuestionList,
    crossTabTitle: "Highest & Lowest Rated Statements",
    crossTabRows,
    keywords,
    questionCount: (questions || []).length,
  };
}

// Fetches and computes live analysis for an admin-built survey (one that
// collects real responses via QR code / shareable link), refetching whenever
// a different survey is selected.
// Exports a survey's analysis (per-question average scores and response
// counts) as a CSV file, generated and downloaded entirely client-side.
function downloadSurveyAnalysisCSV(survey, analysis) {
  const rows = [["Question", "Average Score (out of 5)", "Responses"]];
  const questionList = analysis?.fullQuestionList || analysis?.ratingData || [];
  questionList.forEach((q) => {
    rows.push([q.statement || q.fullQuestion || q.question || "", q.avg ?? "", q.n ?? ""]);
  });
  const csv = rows.map((r) => r.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${(survey.name || "survey").replace(/[^a-z0-9]+/gi, "_")}_analysis.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Exports every participant in a project — name, email, cohort, group,
// status, every configured grading criterion's score, and the computed
// weighted total — as a single CSV, generated client-side.
function downloadRosterCSV(projectParticipants, project) {
  const criteria = project?.criteria || [];
  const groupOf = (employeeId) => (project?.groups || []).find((g) => g.participantIds.includes(employeeId));
  const header = ["Name", "Email", "Employee ID", "Cohort", "Department", "Status", "Group", ...criteria.map((c) => `${c.name} (${c.weight}%)`), "Weighted Total"];
  const rows = [header];
  projectParticipants.forEach((p) => {
    const group = groupOf(p.employeeId);
    const criteriaValues = criteria.map((c) => {
      const raw = p.criteriaScores ? p.criteriaScores[c.name] : undefined;
      return raw !== undefined && raw !== null ? raw : "";
    });
    const total = computeWeightedTotal(p, project);
    rows.push([
      p.name || "",
      p.email || "",
      p.employeeId || "",
      p.cohort || "",
      p.dept || "",
      p.status || "",
      group?.name || "",
      ...criteriaValues,
      total !== null ? total : (p.avgScore ?? ""),
    ]);
  });
  const csv = rows.map((r) => r.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${(project?.name || "roster").replace(/[^a-z0-9]+/gi, "_")}_roster.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// PowerPoint export — generates a real, branded .pptx client-side using
// pptxgenjs, with native (editable) charts rather than baked-in images.
// Gradient fills aren't supported by pptxgenjs directly, so the title/
// closing slide backgrounds are rendered on an in-memory canvas (same
// technique as the certificate image) and embedded as a background image.
// ---------------------------------------------------------------------------

const PPTX_NAVY = "0B2559";
const PPTX_NAVY_MID = "0E3F7C";
const PPTX_TEAL = "0F9A8E";
const PPTX_LIKERT_COLORS = ["DC2626", "F97316", "FBBF24", "38BDF8", "0D9488"];

function makeGradientBackgroundDataUrl(width = 1600, height = 900) {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  const gradient = ctx.createLinearGradient(0, 0, width, height);
  gradient.addColorStop(0, "#0b2559");
  gradient.addColorStop(0.55, "#0e3f7c");
  gradient.addColorStop(1, "#0f9a8e");
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, width, height);
  return canvas.toDataURL("image/png");
}

function pptxTitleSlide(pres, bgDataUrl, kicker, title, subtitle, logoDataUrl) {
  const slide = pres.addSlide();
  slide.background = { data: bgDataUrl };
  if (logoDataUrl) {
    // Client logo, if the project has one configured — small white card so
    // logos of any background color stay legible against the dark gradient.
    slide.addShape("roundRect", { x: 0.6, y: 0.45, w: 1.7, h: 0.7, rectRadius: 0.08, fill: { color: "FFFFFF" }, line: { type: "none" } });
    slide.addImage({ data: logoDataUrl, x: 0.75, y: 0.55, w: 1.4, h: 0.5, sizing: { type: "contain", w: 1.4, h: 0.5 } });
  } else {
    slide.addText("MERCER", {
      x: 0.6, y: 0.5, w: 8, h: 0.5,
      fontFace: "Helvetica", fontSize: 16, bold: true, color: "FFFFFF", charSpacing: 3,
    });
  }
  slide.addText(kicker, {
    x: 0.6, y: 2.7, w: 12, h: 0.5,
    fontFace: "Helvetica", fontSize: 15, color: "C8D7EB", charSpacing: 2,
  });
  slide.addText(title, {
    x: 0.6, y: 3.15, w: 12.1, h: 1.6,
    fontFace: "Georgia", fontSize: 38, bold: true, color: "FFFFFF",
  });
  slide.addText(subtitle, {
    x: 0.6, y: 4.85, w: 12, h: 0.5,
    fontFace: "Helvetica", fontSize: 15, color: "D2DCF0",
  });
  return slide;
}

function pptxClosingSlide(pres, bgDataUrl, message) {
  const slide = pres.addSlide();
  slide.background = { data: bgDataUrl };
  slide.addText("MERCER", {
    x: 0.6, y: 3.3, w: 12, h: 0.5,
    fontFace: "Helvetica", fontSize: 16, bold: true, color: "FFFFFF", charSpacing: 3, align: "center",
  });
  slide.addText(message, {
    x: 1, y: 3.75, w: 11.3, h: 0.8,
    fontFace: "Helvetica", fontSize: 18, color: "D2DCF0", align: "center",
  });
  return slide;
}

function pptxSectionHeader(slide, label, title) {
  slide.addText(label, {
    x: 0.6, y: 0.45, w: 10, h: 0.35,
    fontFace: "Helvetica", fontSize: 13, bold: true, color: PPTX_TEAL, charSpacing: 2,
  });
  slide.addText(title, {
    x: 0.6, y: 0.78, w: 12, h: 0.6,
    fontFace: "Georgia", fontSize: 26, bold: true, color: PPTX_NAVY,
  });
}

function pptxPageNumber(slide, current, total, color) {
  slide.addText(`${current} / ${total}`, {
    x: 12.1, y: 7.05, w: 1, h: 0.35,
    fontFace: "Helvetica", fontSize: 10, color: color || "94A3B8", align: "right",
  });
}

async function generateSurveyPPTX(survey, analysis, logoDataUrl) {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE"; // 13.3" x 7.5"
  const bgDataUrl = makeGradientBackgroundDataUrl();
  const dateStr = new Date().toISOString().slice(0, 10);

  pptxTitleSlide(pres, bgDataUrl, "SURVEY ANALYSIS REPORT", survey.name, `Generated ${dateStr}${analysis?.responses ? ` · ${analysis.responses} responses` : ""}`, logoDataUrl);

  // Key metrics
  const metricsSlide = pres.addSlide();
  pptxSectionHeader(metricsSlide, "OVERVIEW", "Key Metrics");
  const avgOverall = analysis?.ratingData?.length
    ? (analysis.ratingData.reduce((s, r) => s + r.avg, 0) / analysis.ratingData.length).toFixed(1)
    : "—";
  const metrics = [
    { label: "Total Responses", value: String(analysis?.responses ?? "—") },
    { label: "Average Score", value: `${avgOverall} / 5` },
    { label: "Questions Rated", value: String(analysis?.ratingData?.length ?? 0) },
  ];
  metrics.forEach((m, i) => {
    const x = 0.6 + i * 4.15;
    metricsSlide.addShape("roundRect", { x, y: 1.9, w: 3.85, h: 2.1, rectRadius: 0.12, fill: { color: "F4F7FB" }, line: { type: "none" } });
    metricsSlide.addText(m.value, { x, y: 2.15, w: 3.85, h: 1, align: "center", fontFace: "Georgia", fontSize: 40, bold: true, color: PPTX_NAVY });
    metricsSlide.addText(m.label, { x, y: 3.2, w: 3.85, h: 0.5, align: "center", fontFace: "Helvetica", fontSize: 14, color: "64748B" });
  });
  if (analysis?.choiceSubtitle) {
    metricsSlide.addText(analysis.choiceSubtitle, { x: 0.6, y: 4.4, w: 12, h: 0.5, fontFace: "Helvetica", fontSize: 13, color: "94A3B8" });
  }

  // Rating questions — native horizontal bar chart, top 10 by score
  if (analysis?.ratingData?.length) {
    const chartSlide = pres.addSlide();
    pptxSectionHeader(chartSlide, "RATINGS", "Average Score by Question");
    const top = analysis.ratingData.slice(0, 10);
    chartSlide.addChart(pres.ChartType.bar, [{
      name: "Average Score",
      labels: top.map((r) => r.question),
      values: top.map((r) => r.avg),
    }], {
      x: 0.6, y: 1.6, w: 12.1, h: 5.4,
      barDir: "bar",
      showTitle: false,
      showLegend: false,
      showValue: true,
      dataLabelPosition: "outEnd",
      dataLabelColor: PPTX_NAVY,
      dataLabelFontSize: 10,
      dataLabelFormatCode: "0.0",
      chartColors: [PPTX_TEAL],
      catAxisLabelColor: "334155",
      catAxisLabelFontSize: 10,
      valAxisLabelColor: "94A3B8",
      valAxisMinVal: 0,
      valAxisMaxVal: 5,
      valGridLine: { color: "E2E8F0", size: 1 },
      catGridLine: { style: "none" },
    });
  }

  // Sentiment breakdown — native pie chart. Colors come from each data
  // item's own assigned color (already correctly matched to its label
  // upstream) rather than a separate hardcoded array, which previously had
  // the wrong order and showed "Strongly Agree" in red.
  if (analysis?.choiceData?.length) {
    const pieSlide = pres.addSlide();
    pptxSectionHeader(pieSlide, "SENTIMENT", analysis.choiceTitle || "Overall Sentiment");
    const pieColors = analysis.choiceData.map((d, i) => (d.color ? d.color.replace("#", "") : PPTX_LIKERT_COLORS[i % PPTX_LIKERT_COLORS.length]));
    pieSlide.addChart(pres.ChartType.pie, [{
      name: "Sentiment",
      labels: analysis.choiceData.map((d) => d.name),
      values: analysis.choiceData.map((d) => d.value),
    }], {
      x: 2.6, y: 1.6, w: 8, h: 5.2,
      showTitle: false,
      showLegend: true,
      legendPos: "b",
      legendColor: "334155",
      showValue: true,
      dataLabelColor: "FFFFFF",
      dataLabelFontSize: 11,
      chartColors: pieColors,
    });
  }

  // Key themes — chip layout, sized by weight
  if (analysis?.keywords?.length) {
    const themeSlide = pres.addSlide();
    pptxSectionHeader(themeSlide, "OPEN FEEDBACK", "Key Themes");
    let x = 0.6, y = 1.9;
    const rowH = 0.85;
    analysis.keywords.forEach((k) => {
      const fontSize = 14 + k.weight * 3;
      const w = Math.min(5.5, 1.2 + k.word.length * (fontSize / 55));
      if (x + w > 12.7) { x = 0.6; y += rowH + 0.25; }
      themeSlide.addShape("roundRect", { x, y, w, h: rowH, rectRadius: rowH / 2, fill: { color: "EAF6F4" }, line: { type: "none" } });
      themeSlide.addText(k.word, { x, y, w, h: rowH, align: "center", valign: "middle", fontFace: "Helvetica", fontSize, bold: true, color: PPTX_NAVY_MID });
      x += w + 0.3;
    });
  }

  pptxClosingSlide(pres, bgDataUrl, "Thank you");
  await pres.writeFile({ fileName: `${(survey.name || "survey").replace(/[^a-z0-9]+/gi, "_")}_analysis.pptx` });
}

async function generateComparisonPPTX(surveyA, analysisA, surveyB, analysisB, logoDataUrl) {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE";
  const bgDataUrl = makeGradientBackgroundDataUrl();
  const dateStr = new Date().toISOString().slice(0, 10);

  const titleSlide = pptxTitleSlide(pres, bgDataUrl, "PRE & POST IMPACT REPORT", `${surveyA.name} vs ${surveyB.name}`, `Generated ${dateStr}`, logoDataUrl);
  titleSlide.addNotes(`This report compares "${surveyA.name}" (pre) against "${surveyB.name}" (post) across every behavioral statement common to both surveys.`);

  const matches = matchQuestionsAcrossSurveys(analysisA, analysisB);
  const increases = matches.filter((m) => m.band === "increase");
  const decreases = matches.filter((m) => m.band === "decrease");
  const same = matches.filter((m) => m.band === "same");
  const { onlyInA, onlyInB } = countUnmatchedStatements(analysisA, analysisB);
  const overallPre = matches.length ? Math.round(matches.reduce((s, m) => s + m.pre, 0) / matches.length) : null;
  const overallPost = matches.length ? Math.round(matches.reduce((s, m) => s + m.post, 0) / matches.length) : null;

  // Table slides are numbered on their own sequence (title/summary/closing
  // aren't part of "page X of Y" here, since that count matters to someone
  // flipping through the detailed tables specifically).
  const ROWS_PER_SLIDE = 9;
  const totalTableSlides = Math.max(1, Math.ceil(matches.length / ROWS_PER_SLIDE));

  // Executive summary — headline before/after number, then narrative bullets
  // and a donut chart giving the improved/declined/unchanged split, so the
  // top-line result and the shape of the detail are both visible at once.
  const summarySlide = pres.addSlide();
  pptxSectionHeader(summarySlide, "OVERVIEW", "Executive Summary");
  let bodyTop = 1.7;
  if (overallPre !== null) {
    summarySlide.addShape("roundRect", { x: 0.6, y: 1.55, w: 12.1, h: 1, rectRadius: 0.1, fill: { color: "F4F7FB" }, line: { type: "none" } });
    summarySlide.addText([
      { text: "Overall Agreement  ", options: { fontSize: 13, color: "64748B" } },
      { text: `${overallPre}%`, options: { fontSize: 24, bold: true, color: PPTX_NAVY } },
      { text: "  →  ", options: { fontSize: 18, color: "94A3B8" } },
      { text: `${overallPost}%`, options: { fontSize: 24, bold: true, color: PPTX_TEAL } },
      { text: `   ${overallPost >= overallPre ? "▲" : "▼"} ${Math.abs(overallPost - overallPre)} pts`, options: { fontSize: 15, bold: true, color: overallPost >= overallPre ? PPTX_TEAL : "DC2626" } },
    ], { x: 0.9, y: 1.55, w: 11.5, h: 1, valign: "middle" });
    bodyTop = 2.85;
  }
  const summaryLines = [];
  if (matches.length > 0) {
    summaryLines.push({
      text: `Comparing "${surveyA.name}" and "${surveyB.name}" across ${matches.length} matched statement${matches.length === 1 ? "" : "s"}: ${increases.length} improved, ${decreases.length} declined, and ${same.length} stayed the same.`,
      options: { bullet: false, paraSpaceAfter: 14 },
    });
    const topGain = increases[0];
    const topDrop = decreases[decreases.length - 1];
    if (topGain) {
      summaryLines.push({ text: `Largest improvement: "${topGain.statement}" — ${topGain.pre}% → ${topGain.post}% (+${topGain.impact} pts).`, options: { bullet: true, paraSpaceAfter: 10, breakLine: true } });
    }
    if (topDrop && topDrop.impact < 0) {
      summaryLines.push({ text: `Largest decline: "${topDrop.statement}" — ${topDrop.pre}% → ${topDrop.post}% (${topDrop.impact} pts).`, options: { bullet: true, paraSpaceAfter: 10, breakLine: true } });
    }
    summaryLines.push({ text: `${same.length} statement${same.length === 1 ? "" : "s"} showed no change between the two surveys.`, options: { bullet: true, breakLine: true, paraSpaceAfter: onlyInA + onlyInB > 0 ? 10 : 0 } });
    if (onlyInA + onlyInB > 0) {
      summaryLines.push({
        text: `${onlyInA + onlyInB} statement${onlyInA + onlyInB === 1 ? "" : "s"} couldn't be matched between the two surveys (question wording differs) and ${onlyInA + onlyInB === 1 ? "is" : "are"} not included in the tables below.`,
        options: { bullet: true, breakLine: true, italic: true, color: "94A3B8", fontSize: 13 },
      });
    }
    summarySlide.addText(summaryLines, { x: 0.6, y: bodyTop, w: 7.6, h: 7.3 - bodyTop, fontFace: "Helvetica", fontSize: 16, color: "334155", valign: "top", lineSpacingMultiple: 1.3 });

    summarySlide.addChart(pres.ChartType.doughnut, [{
      name: "Statements",
      labels: ["Improved", "Declined", "No Change"],
      values: [increases.length, decreases.length, same.length],
    }], {
      x: 8.5, y: bodyTop, w: 4.2, h: 7.1 - bodyTop,
      showTitle: false,
      showLegend: true,
      legendPos: "b",
      legendColor: "334155",
      legendFontSize: 11,
      showValue: true,
      dataLabelColor: "FFFFFF",
      dataLabelFontSize: 12,
      chartColors: ["0D9488", "DC2626", "94A3B8"],
      holeSize: 60,
    });
  } else {
    summaryLines.push({ text: "No matching statements were found between these two surveys — question wording may differ, or they may not share a common question set.", options: { bullet: false } });
    summarySlide.addText(summaryLines, { x: 0.6, y: bodyTop, w: 12.1, h: 7.3 - bodyTop, fontFace: "Helvetica", fontSize: 16, color: "334155", valign: "top", lineSpacingMultiple: 1.3 });
  }
  summarySlide.addNotes(`Headline: overall agreement moved from ${overallPre ?? "—"}% to ${overallPost ?? "—"}%. Use this slide to set up the narrative before the statement-by-statement detail that follows.`);

  // Pre/Post/Impact tables, paginated ~9 rows per slide (matches the
  // reference deck's density) with a color legend and sample sizes on the
  // first table slide, a page indicator on every table slide, and
  // alternating row shading so long tables stay easy to scan.
  const headerRow = [
    { text: "Pre", options: { bold: true, color: "FFFFFF", fill: { color: PPTX_NAVY }, align: "center", fontSize: 11 } },
    { text: "Behavioral Statement", options: { bold: true, color: "FFFFFF", fill: { color: PPTX_NAVY }, align: "left", fontSize: 11 } },
    { text: "Post", options: { bold: true, color: "FFFFFF", fill: { color: PPTX_NAVY }, align: "center", fontSize: 11 } },
    { text: "Impact", options: { bold: true, color: "FFFFFF", fill: { color: PPTX_NAVY }, align: "center", fontSize: 11 } },
  ];
  const bandColor = (band) => (band === "increase" ? "0D9488" : band === "decrease" ? "DC2626" : "94A3B8");
  const bandArrow = (band) => (band === "increase" ? "▲" : band === "decrease" ? "▼" : "—");

  let tableSlideIndex = 0;
  for (let i = 0; i < matches.length; i += ROWS_PER_SLIDE) {
    tableSlideIndex++;
    const chunk = matches.slice(i, i + ROWS_PER_SLIDE);
    const tableSlide = pres.addSlide();
    pptxSectionHeader(tableSlide, "BY STATEMENT", "Pre & Post Impact");
    const nNote = (analysisA?.responses || analysisB?.responses)
      ? `% of participants who selected Agree or Strongly Agree · N=${analysisA?.responses ?? "?"} (pre), N=${analysisB?.responses ?? "?"} (post)`
      : "% of participants who selected Agree or Strongly Agree";
    tableSlide.addText(nNote, { x: 0.6, y: 1.35, w: 8, h: 0.35, fontFace: "Helvetica", fontSize: 12, color: "94A3B8" });
    pptxPageNumber(tableSlide, tableSlideIndex, totalTableSlides, "94A3B8");
    tableSlide.addNotes(`Statements ${i + 1}–${i + chunk.length} of ${matches.length}, sorted by largest improvement first. Walk through any standout rows rather than reading every line.`);

    if (i === 0) {
      // Legend, once, on the first table slide
      const legendY = 1.35;
      [["increase", "Increase"], ["decrease", "Decrease"], ["same", "No Change"]].forEach(([band, label], idx) => {
        const lx = 8.6 + idx * 1.5;
        tableSlide.addText(bandArrow(band), { x: lx, y: legendY, w: 0.3, h: 0.3, fontSize: 12, bold: true, color: bandColor(band) });
        tableSlide.addText(label, { x: lx + 0.3, y: legendY, w: 1.2, h: 0.3, fontSize: 10, color: "64748B" });
      });
    }

    const rows = [headerRow, ...chunk.map((m, rowIdx) => {
      const rowFill = rowIdx % 2 === 1 ? { color: "F8FAFC" } : undefined;
      const cellOpts = (extra) => ({ ...extra, ...(rowFill ? { fill: rowFill } : {}) });
      return [
        { text: `${m.pre}%`, options: cellOpts({ align: "center", fontSize: 12, color: "334155", valign: "middle" }) },
        { text: m.statement, options: cellOpts({ align: "left", fontSize: 11, color: "334155", valign: "middle" }) },
        { text: `${m.post}%`, options: cellOpts({ align: "center", fontSize: 12, color: "334155", valign: "middle" }) },
        { text: `${bandArrow(m.band)} ${m.impact > 0 ? "+" : ""}${m.impact}%`, options: cellOpts({ align: "center", fontSize: 12, bold: true, color: bandColor(m.band), valign: "middle" }) },
      ];
    })];

    tableSlide.addTable(rows, {
      x: 0.6, y: 1.85, w: 12.1,
      colW: [1.1, 8.3, 1.1, 1.6],
      border: { type: "solid", color: "E2E8F0", pt: 0.5 },
      autoPage: false,
      rowH: 0.5,
    });
  }

  const closing = pptxClosingSlide(pres, bgDataUrl, "Thank you");
  closing.addNotes("Close with the headline result and invite questions.");
  const fileSafe = `${(surveyA.name || "survey")}_vs_${(surveyB.name || "survey")}`.replace(/[^a-z0-9]+/gi, "_");
  await pres.writeFile({ fileName: `${fileSafe}_comparison.pptx` });
}

function useLiveSurveyAnalysis(survey) {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!survey || survey.sourceType !== "built") {
      setAnalysis(null);
      return;
    }
    let cancelled = false;
    setLoading(true);
    (async () => {
      try {
        const res = await fetch(`/api/survey-response?surveyId=${encodeURIComponent(survey.id)}`);
        const data = res.ok ? await res.json() : { responses: [] };
        if (!cancelled) setAnalysis(computeAnalysisFromResponses(survey.questions, data.responses || []));
      } catch (err) {
        if (!cancelled) setAnalysis(computeAnalysisFromResponses(survey.questions, []));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [survey?.id, survey?.sourceType]);

  return { analysis, loading };
}

// Read-only fetch of the shared survey list, for the participant-facing
// Surveys section and the notification bell. Admin's SurveyAnalytics keeps
// its own separate read+write copy for managing surveys.
function useSharedSurveys() {
  const [surveys, setSurveys] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(`/api/data?key=${encodeURIComponent(SURVEYS_STORAGE_KEY)}`);
        if (res.ok) {
          const data = await res.json();
          const parsed = JSON.parse(data.value || "[]");
          if (!cancelled && Array.isArray(parsed)) setSurveys(parsed);
        }
      } catch (err) {
        // nothing saved yet
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  return { surveys, loading };
}

// Read-only fetch of the shared meetings list, for the participant-facing
// Meetings tab. Admin's AttendancePanel keeps its own separate read+write copy.
function useSharedMeetings() {
  const [meetings, setMeetings] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(`/api/data?key=${encodeURIComponent(ATTENDANCE_STORAGE_KEY)}`);
        if (res.ok) {
          const data = await res.json();
          const parsed = JSON.parse(data.value || "[]");
          if (!cancelled && Array.isArray(parsed)) setMeetings(parsed);
        }
      } catch (err) {
        // nothing saved yet
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  return { meetings, loading };
}

// Checks whether a specific participant has already submitted a response to
// a given survey, so the Surveys section can show "Completed" instead of
// letting them answer twice.
function useHasResponded(surveyId, employeeId) {
  const [hasResponded, setHasResponded] = useState(false);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!surveyId || !employeeId) { setLoading(false); return; }
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(`/api/survey-response?surveyId=${encodeURIComponent(surveyId)}`);
        const data = res.ok ? await res.json() : { responses: [] };
        const found = (data.responses || []).some((r) => r.respondentEmployeeId === employeeId);
        if (!cancelled) setHasResponded(found);
      } catch (err) {
        // assume not responded on error
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [surveyId, employeeId]);

  return { hasResponded, loading };
}

// Fetches live check-in data for an attendance session, refetching whenever
// requested (exposes a manual refresh since check-ins arrive continuously
// while an admin is watching the panel, unlike survey analysis which is
// viewed after the fact).
function useAttendanceCheckins(sessionId) {
  const [checkins, setCheckins] = useState([]);
  const [loading, setLoading] = useState(false);

  const refresh = React.useCallback(async () => {
    if (!sessionId) { setCheckins([]); return; }
    setLoading(true);
    try {
      const res = await fetch(`/api/attendance?sessionId=${encodeURIComponent(sessionId)}`);
      const data = res.ok ? await res.json() : { checkins: [] };
      setCheckins(data.checkins || []);
    } catch (err) {
      setCheckins([]);
    } finally {
      setLoading(false);
    }
  }, [sessionId]);

  useEffect(() => { refresh(); }, [refresh]);

  return { checkins, loading, refresh };
}

const gradingCriteriaDefault = [
  { id: "1", name: "Innovation", weight: 25 },
  { id: "2", name: "Critical Thinking", weight: 20 },
  { id: "3", name: "Clarity", weight: 20 },
  { id: "4", name: "Depth of Analysis", weight: 25 },
  { id: "5", name: "Grammar", weight: 10 },
];

const recentUploads = [
  { name: "cohort_b_reflections_march.docx", type: "docx", size: "1.2 MB", when: "2026-03-10 · 9:14 AM", status: "Processed" },
  { name: "participant_roster_q1.xlsx", type: "xlsx", size: "84 KB", when: "2026-03-09 · 4:02 PM", status: "Processed" },
  { name: "strategic_workforce_planning_sramirez.pdf", type: "pdf", size: "310 KB", when: "2026-03-10 · 9:16 AM", status: "Grading" },
];



// This environment's Tailwind has no JIT compiler, so arbitrary-value classes
// like `from-[#0b2559]` are not reliably rendered. Brand colors are applied
// via inline styles instead, kept here in one place for easy editing.
const BRAND_NAVY = "#0b2559";
const BRAND_NAVY_MID = "#0e3f7c";
const BRAND_TEAL = "#0f9a8e";
const brandGradient = (dir = "to bottom right") => ({
  backgroundImage: `linear-gradient(${dir}, ${BRAND_NAVY} 0%, ${BRAND_NAVY_MID} 55%, ${BRAND_TEAL} 100%)`,
});

// User-supplied real Mercer logo asset, embedded as a data URI so the
// artifact is self-contained (no external file hosting available at runtime).
const MERCER_LOGO_DATA_URL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAIAAAAiOjnJAAAQAElEQVR4Aex9CbhlRXXuv6r23uecO3SDOESIkjjlg/hi8l6++CX5YpKX5CVRwAEIxhejSZ550ScqM4j6jOFFDSgKNDM00HMDDQQBAZmJyCAyQ0dmoZmb7r7DOXuoqvfX3uece7v7Nn27ufcM9+79rVOnqnbtqlV//bVqVe2+p9VTDVdKicBOI/BEwz1ed4833BOxe7Lunhh3T9fds4lTmzaZDSNpGZYI7BQCZmQEm0bdyAgjduMow2zTKMUp48Q6KcMSgZ1DwAm8QDkvsEJRVkRBK9FBGZYI7BQCiqyyyLklDBVEHACBMoBxKMMSgZ1DILOw8CuecS4XJhkBKQbyq5QSgZ1DQAVQJJGGaFGMedHQUJaMK2UeIDBbA21gDVwLQOcvn1QiKKVE4PUjoJCbLmkySkm5EpYIvA4EaKs8h+ixGyjnhQ67j6C8SgReNwJkUruOwlTR13IKpZQI7BQCuYlqs6qglDgnzG8TrYyUCMwgAspCSikR2EkEBDaXgpF0tLwIT0pBX77ILMNOIzCX2iO9mt3hKpjH/ProWcZT+VJKBHYCgdw++Rc7klsv1iClxcrnVhnMCALk1uR6/BlE05kHykiJwI4igHy94wJI4fJH14oRCiMorxKBGUOgZbhKYs0YpGVFkxEoiTUZjTI+YwiUxJqAsozNIAIlsWYQzLKqCQRKYk1gUcZmEIGSWDMIZlnVBAIlsSawKGMziEBJrBkEs6xqAoGSWBNYlLEZRKCniTWD/Syr6jACJbE6DPh8aa4k1nwZ6Q73syRWhwGfL811i1huvgA8X/s568Qig7zk/6rQFX/AAef/nT3DpuwQ9g7Np9qRHXq87wv7P6ly/s/3VBuAIjKpZ7zVlknZHY3OOrG22Zvibzj8bQLjv3b4wxp2+JnygQ4hsGPE2mmlxCEXYegraf8Rm0/s0Kf4F475I6wk/55XgYVYyf91eYEEQ/a/CFu2vF2AEd7sisw6sdhlyqS+bZ5q3nD5AtdMlF87hkATUWK4Y8/NaulZJ9ZU2hOJtrTui825NU102o8z0qphnn4TsVwIIKVnQJh1Ym3Pi5zEDOJC6Rloel6RSXzqPdxmnVjt4dkOw+alw9QGZ/qRnE1w4mXLpzyGnKjej20W27JE59KzTizf0e10p1WEuFC2U7i8Df9nesivyYbKQ0ck8/x2sFVG+85sR2adWJM7wE0KpcgJAhS/MReGMImBQaTFZY6ZWkMEaWqcg/81HGdb7hfnYfF0Hm6eyrPmfkBkiAnDRqOhiRS4S3TKKeWkEnic/O/rOY9bkqYshp27XvdTHSXWZG3HxlICVKmgPo4g0GEoTGqttIIxZJUdHtbWeu4EgTLMmvzwPI5z4hEMhsPDNfLGZC4IAq0kyxA3ECpPqSiA1qjVQsep2SWsZp9Y5AZlonuu8A+GhkLmpQ1/iDw+kr2wbvTxn79sEtgUtN9KVByDuBDEJMmUmkpPlmMV80zIGPKJblSWuCw2xCWLkSWgyX91feOVl9MNr8aEbtPGRpZZka5hRMVmd2QKGrEN6/tYUIyhazSsZ4vDiy+u//a3Tjj0i0d887gTfnzL3aHnG20YTIZq1TNKRIeciayiFK51jshw5QOtuNY6CgMiFmgsueDCr3z5//6fz37h3LPOSxt2l4XVasTJWecTXYFt1om1Ra+c9zctQFYpTietEY+nax9+9KH71/746htPOP77Tz7xinPgnIsi/2iWuSiSJHE+4T+MUHxsfn44Pb1Y1LjgWTTGrVjccdsjK5dd9JObf/LQPQ88/viTSZKOjtpGnA7UKt1CqUPEKsyVZxVhyIXmKk29Q9BoJElsTCYIB5975qXFi5ekibdYhdNQrIYMuwVQr7VL76pSEWOsy01XEKgXXqifccY5L657ETaAqqQNI9Ch9iNrWRrduXzzs9oyKUXZugn2uaDLwMDQ4OAwnAbCtJFcdeXV11xzPWlH4BhyNWSkUlFb1zA/c4zxm2URcgthKHStfviDa+656164QMIqEJBLLMOlIAzntvM+Mf404RMJpVS1Ko1xbgDj+ihdAcUjP5EgHm2c8O3vvfTcGN0qUoowWeviOGs9yUoordT8+W71tFLRjUZG90ABLsPDDz9x4oknpUkG0Y6nNtZGvOe474EIw665DVSvpfLsfCvn932tujVc0BQgSRBV/GpIew5noAiEIAk2vRwvOunMxhgCxSMuG4QIyDJfRU4pYkXxSX4mYkz0p7ALlEm6t1JkRiFOXC6w4Dx0YejNUhLjpZfGT/ne6dlIrCREsS44HgRmrIsPZqmDcB1gqgsy68Taqk85ObbI9V6Xy/OoT4gMN91wy1VX/EgcaadCf0jjwfIFilI+Nhc/7B2l1TPHC8X/oNXMEiEgQg+B5wtc7C44d+mdd9wdDS20fnND6EArxg/F8tNVKbTpqgpbNK7olJqN6zeeddbip596qVLhCaqpBBHgGdk8vPBR5NdELE/OiWCrPnGCFcLRIqt4tq4V7rj9/tWrL8oaPGDPN0E91nWq2mMaOafJJqhnn3x68dkXcEF0RvxLnUJNgk4p4nMvZNcozX41bZc3W80c/5WloGv18stjixad1mjEKqryPRhoyvzNiU/Xx7XrCkxgUcSU0ibN0EiUrl55xbUXX3T1LgsVX1bkd4k1xTih0O0AE3n+HA3oIYCcmcQ1CzpNJsN5i5fef+9D9KtsagAOIgVbX7J1VqdyplaoU61v3Q59A0sHVVcHbT3ly4ol56146P71w4PNHYDAIod76yf7OWdrArituiPOcXODgRpuu+3ulSsuCnSIOF8EuTpuVbrrGb1GLNgsCyoV4R4nqJlUnn3mxdNOPXP9K4VXygGwwnVRLPy/jegqejPWODs1RV3FCij5pRnafEIZ/HztBnoI2VicjSfh0C507qOoyuk2RRVdzeo5YgVBkKUpJZAAmbMNc+tNP7nlptvovdKtV44KU7qKWWca95Mnb4ldLuwXQ4s1F//7vXffC54j6ChtJDxjTxoNpblC5oV7Jui5QeKpMd0KpXmorKA8t0xmzzj97IcefIqHpbWajuNUScAJTaph6tneM+hOQ5F8tmxWrrnzBU8WFA/WWcAZaEEU4Lpr/mPVigt5tAyyzTiledInLGcNPS1MXG1STmR1OqY63eD22tP5Rd6QYcJ5qYLGeGPdL9adv3gpIR7ZYBYODcT1RGseEmYi26uuL+/TNEFEUp6kAwVnlOCZpzcuWbKi0Ui8w27ZMeWccLKBTJQCCJ/LG70gPUesnE/Cyzoj4jTnqZNkdPy6H9145RU3BoFu1BFFFXEgA63HtRdgnHkdwkCz0mpFRRHIqriOyy6+/IG77rZ8u0XzJIGIN9ve1SS/pCAWn6B4XvKLQ8upyEhXhK13pd1tNapcxs2PUrleLndKJQgQVrLx+PzFy3/+yDMEKxC+DjJ6MzS3VWG/5ieJnzXj4y5NYDP88KqbV6++GBLBiRAeFYiQTAo+BDjPiuVPuOvB5heLbZ7RkZTqSCs70kgQEht67TxgcM6YLCG0RBJOP/rzp849Z2l9HNwXAmKtCzRRa05Q9PPF2VKo79ihPKa1GhwMjHHVKl54fnTFilWvrh8JK4NKhcL3zY7vBC23xhAHBee2BIHZeTVdC3qOWEprv77RsxCnyRsi5JzNCF9k69n1191y6ZorQ41aRKNm3WRAu4bhrDQcx5l/SR+qsTFccMHSB+990BmXjjec96ssadVeBEUEjqsj1bD89Ij0HLGszScfwQKEpwv0NJQSkTCoILPJeLJ0yYqf3vUEQae/5fiCA31/tYxUuyPM4PofaOU991tuvuPyy68Qvi11IgTBWJo3TjfkOOXPbMGnLZJ5kY4HPUcsmEyUUvSrHGmT2iyjoafxT5NEVQa4/V637vkVy1a++uII8QUC5SGbwNin+u0zSXtGKX5943LPOfX4Y+tWLFs98vIGF3ML7L11JzxeUFprggQQG1+ekUk8y1Mkp//u2icfl661PlXDWpzNLJdDIbAh6K4qnizYMIrs+Gh1oIrMXPXvV1571U1IEIEAC1/9iKafz2egePLFF4l0Owh6S4pmmEcp4r0T0qGERmo5g1IdiKaKxgSCaoikjtXLLvnpbfciJZOqnF1++ROhA0rxD4qnj/fX80irUxNj6su0cjv8PaFEhxveRnOFGS9CFqF6TaHFklqtsXFjoKNA185fvOTRtRv5MtE5nj5EScKhsVHE3aINOCx8tE8k5STyXVBhGDYaDWqtoWmmyaq77lh75eXXKKlEg7uYeszJpsMQ5BKF5SYJXYY8RVZOWK42iPmtTgeq0w2+jvYcTVKlwrc9WZa98Nxzp5xyytgo6OXDqVolDFTA8RARHYCj5dvhABTiEz36od5KkGuOUGnlRItf4J9ft/HE735vfHQ045FokkRDQxgfV6pvxqtvFOUqUKkOcM1QJA5UGNVuuP6mH/zgSvoiWeZEQEkaaRhKmjLuIFsxyQkoW2V3N4NqU9I0zk1vmMRGKdTHwfOFR+69X9NnZ5qnLlkWLFiQ1uvd1Xb6rc8msaavxTRKShDFcQzRtE9wSBsJ7dfixef9fO1ToQgXDiaDIBDH2Z9EkQYLwYnzMo3qu1YkyywPqyqVCjUwGUQ0p8qtt9y+fNmqYHDYxCmgJPBuu7UZaM1Yrh+kb4jlwUyyysCQTTJYEZ6jZuaFXzy/+NwlGzcmUQhSKVTClTEKtTF0MHIRhv7Rnv0EgeLKHgTc39Eoo1bFU09uWHLBCpv/w5hogEY61fll6/UgDHu2I1so1jfEcnFaXbAg5lqglA4rzgqgVaV29RU/+sFlV41uREBbxndAlseqOktjIfuwGav4BGWL/nc9WVDF0VZZcG68+Fxy4ao19975s3BwAc9WbOYQVIQ7wbGxcGjIn5l2XePpKdA3xNLVamN0PODGz9JhN54zTrg4QELO77VrH+ciyM4Iz1eNDf1wkVW5eKPlPBoMKD7WQ59Gg06VoqGFgzW4776HVixfDYnSeipRNWskolTK3WKlkqapomXrId1fSxWOxWvd7p17hiTKVw14cBVU4MUpJOaFZ184+8xzNrzqjZZSis6W6j0COecXu63xDLUWK4FCqPHKS43vf/fkLLYwAgmcAXLvClojv3jikH/3QaB6TMcd10fCKBy47ZafrLno0rERmNRppUQ4EqyK0kP9E5lCGeaFoarXHXeCp5565uOPPy1QUAEYorhod4tIP4W9BX0LubZWxJTSyp7iWyFzyaZRa9yFq9esfeSJkPtChzShyWIlhUzxWOezpmQV1eAJFmVwQK677sfX/ehGGmCbWjRSGt1c/JtBFgPXflGg5IkeC6ZQR02R11dZEkQS1bSKfvHkL0477YyxEf9bZFpzRdTgQlkIxPeJAcXHeugjgsEBPP7Yq2ecdubGVzaooKp0KPnpgycTNRVOLQpj/SQ9SCyqRJkMImHdlgCGb3KqZjyGcT/+j9svu/QKhCce6AAAEABJREFUZ2FZHFD+OFRQsAo9elFbbmmXLVn55NrHwIOUhC8KA7FOnBWaXc+qSZpznkxK9XJ0iyHsZVWn1M06a+PRuq7VYCVU4bJlK+668xHFIZmyeK9lOmjBpZde+8OrrlaVIUggmi8KM8u3B1TVs8o27Rb4HrCfBqsvdC3AnTJEpVoF1z0LFUbpptFnn1l33nkXvPzyKHfvHJoeF24V1z7yzPJlKzeu32AbMdgLxVfoKqpUe1zz7arXF8R6rV7wdIcrnzP+NAiVmnLBLdfdfMP1t/h1xHE1fK1nO37PtVukTaU9YrhyxZrHH3hENJkkOopMwn2HTflyB/ArefuBfov0ILEsOHM3w5FKbkvgD9MDTaNleEjtxKYuDAdO/t5pd/5kLQyUgKeOdJB1iNSkxh8NbVb1aybIA8prFtnqZr5iNZ/iV1Po6Sn44ynh6W7GN5kKMAkuXnnVFZdepSsLXOrglMkynmg55bwIyysLLwCLs6WtkWFmj0qhcY8qNx21RMQXc7yEY8NpnqYuHo9Xrbxo40YT11GL0GikdOd5HF+paF+4sx8rpAj5ZDT9J5MEQZAlIOPXPvTkhasv5cGbzVhCwLM3EarGnvAslU8xDnCAKHl0B4LuF+1LpSfDJuLAgaO3AogKAI3MJo3kuutuWLFiFW0Z/eCQb6w524E45mE2pn0JZm5HKcLaEOhIKdIGoyNYterC/7znPmMM9x8+i43lZcAr7w6/+1f6nlh+fhN+8ZdSSsgtTn2nXWYvWrXmkUd+wdeGoQLpx9cmzcIsP6sy2TnydAKbV/kOg4YzbYDm6tprb7j6hz9C5DeznlXiy02olyfRz5fqZ+W97o5uE1mTjwSHzXFQSSzSC8ErL284/bSzX3whyzJkKV18BKq9FNLObSG+thn4UIF2LZ4tIMR00p2xCnQB+RYQ3Amev3hpMhoDAai5Uu0nmtxiZjurPyMTXepP/b3Wkl8cEr+mkFx0T6wGfXkrN9946/KlK6oRQgWOLpcd/0CnP2QwaW3Jn6EhzYlw7tnnP/bwz1VQ5aGuiKb6EPHifMlOazc77c0FYnlKTQyJov8OiA7934oJwjUXX3bPz54S4caR/lfR3ynHj5kUvM7Ls1qadYhzuYAmNZDAZaiP4uof3nT1VddCh5wCDDdTnlq6GdCh2XxXvwqgu6rC62o898l5IpRXIpz9SgnXQW7dk1R05OJ0/XMvnnLSovUvNZI66G/lBTsfSBD45XjtI0+deebZnvqpC4IIfN9MJlHaGs0VbvU7sYCWeQDoBBfzv5j0wrdtoIejw3vveWDNmksd3yHyZEtx589iorXwi5syER/n47nwWUoe3alAhGo0hRWEofDsoxoiaXi7xUXwmafWgSu1BIasUgp8gMKibdki2c7vq4jqK223pyyXnElFhHEaA6fSRrp86ar77nmU/g1PTQEJQ4yPJ0qpWi3g2X3qfyIWM3VlmYkiZFkmIsagWq02GohCLF92yV13/oxHuCDdoXiuRo7PVKO9Vs/cIlaOrnLWi9YuSaoDg4jTrB6/+vwri046dcMrqWaPnTAUYcyPLO1WGOr80RkIrMXAgK7XzcBA4DJHv4rHHNw6PPzAsxdfeOmGdS95c0WLxSNbFtUz1u4MqD6jVXhwZ7TCDldW6F+EedM0WrlYk1aHBxsbN6pKRYVVHVV/eufdq1ZdPDbmPa16HZUooDlLEhuGkUj+7EwESsEfyYY6y/xOsFKhabQjI3b58tVPPfmMVAa9g8WG/Pl7knObiTkok4akX3vHLlAK7enLN0UJbOrHVquQxsmMx5Vo8IJzl9x8021cp2gsyCqWUf7iO58dOpEv2po6zDJ/rMBzjUYjGRzUXASHBtUlay6/5urrEFtxiucgwm0FGWeN2Nflz02tQW/ktoekN9SZKS3EikhSHx3Ydde00TCNVKJKEpvxDZuWLFm27tl6rdpsSZF/FnrmliS2W1QWRRFNF03hk0+8uvSCFfUNI3CaDpau1HjelqYxD9bUnF0JvRfZhLi/vziAFL40pIg3WiaLK4OD+R9Pi/fVnXKZjYYWPHDXz1avvmj9enrWhYPld3D05Weq+9WqbNo0PjQkpNfISH3TpvqiRac9//QzUJGEFahAiWaT9P8Ast8xnKmme6qe/rdYXFwoTVAnuhNWwnh8lI45t4Fah85xOGnDUkh1+QWrH35gLd1qPqTyJzIe04PEbAvv7KTEsR0aGnj1VW45MTxUu+aq6390zQ3QFXCJlICVpuPjUIqeH21baoslmDOBd+aU5Lj2cY/yc3a+niO3OPnzjnCULFRqHMIKIwiDLLNgGR0CoZaBsec2nXbKWa+8WOdLHgWIcGniQ/CXAyg+xg9jbWGyJcxrRfPSjnaf1TObIcQ6uIFKlDXwwL3PnLd4VboxBSIEkaE/LwYR28yszZxoXw3t64QAE9y2QFuQX3yQkkd7PugbRaeBZNGXItxGcSdmLFbV4fvuvn/ZkuWB4g7OhUpoQaZ+oD3kxW3PnSIGIQOaUa5srRgYJ6+EFKqPY9nS1eueXIdwAJmA1GdtvqAl+fw3P8xkOBflNYehDzrMQXIQl2tq83DKwPoyLCZOFX67Ad8h/set92UpuQWXFTWAJocyUcWUA9+iVJtbecVgNiORCrnIMn7llT+89tprkeX/ZrqovghZe7Nagk+LWwgpWUQYMs5C/S3sW393AFuuF+zOlJ2yeUnQpNg4FhXUR+snf59HpqOVEKpwtfgoqUEBKUBhmuU50q0KyRdKkU33jScWvO8jeVYe0FZVIjxw/2NLz1+RbhoLBgctzzba9edl8oB1biF59kTAuxOJvov1t/YtuAvSMGxlgP0qpMjhLavaFLTOJcYYeejeh5dcsMpmZEdRjGTyJR39JK5duRQ3tg49nfiBEX8173vWOWx6FecvXvr0o48DOqs34J0uo5z1Lp3zhwzKqVwYpxTxImSSVVFzhhRGCmGcYuG7wEgfCPXuAy1fU0XCvcV9dqqQdj6TPk5ukQbckamw4hp8oze8euWam298QEsxon45Q9MTAnxK8JoXa5u47/yjocJll11143U3+52g/ycN1pvD3LZNlPS8B7joUiblToo2FZ6U02fR3unA6wBuggqsRDUHzI/ZhBmgheA9ihkfDZW2cRINDCWNLG1k55+39InHX4HjTdLI+pWQUWzfOIhiG5rc4u5O8XHjD8buuP2xlcsusuOpOKW57wvCQBRotPI64Sm1Oeaso3mLM8RLbtsYYbG2FCWYSSnivR5S9V5X8TX1e02gm2PGPlIA50OJIjo9KoiSsQZcAKPuuv3uSy7+d3HeaLX5N6lRmRRvRx35RGFa8vs0SRQYLF+68slHn4KKBMrUE3JNJC+BrS42SWlmW2/uOEMKaWa2v7zm7URfRPpP4y1gDaLQ7+vAUyH2xalAdEALkSn6y8ZYulNKvC8EpXmO5ZSI0L1iJYpJbz8CZ+T8c86/4/YHnUWgtTHOWv+b3qxOckowwvJtcS3fy1mhrSITTAqugMphzUXXXM/jUMcDM82DKhWGbDVNU68MwJCasPJmxGQSiM2SKAos96XOWJdCO74jAOstLKaTdrv9FdkCtP5S3mubjY9Gw4NBRH6kMKnlOzhYCZRtjAfVSIUcKMfXdtaYdGws5EsVt3mXuYSRXk6ddca5Lzy3YXzUDNciw2pSm2WGDSj4EoxMkmaOc84Yz2puAxt13HfPoytXrII/4mcTAVjtpGdEJEsSPkJlAMv9AWguTaprQSMbR9rQkdI0c7zlic0z1dbDE9xita3Mnv/uJ12nBPNNv7z7Jz75iT/+0/8+sHBYD1YQisnqXIX0UCVLxjj0fCoeG5UwqCzcJR4fY3KS2Gbc6TtvuW3Vyoud1XGMWq2ihWZF+dWtWWLLL1KE9sgYOzwEvqQZ2Zicc+bixx9Yiy2I23rOWlRqg9AqqY9mNhXNpjMwpIhBVVlJF+wy9LEDP/zxTxw4tHCY5Gs92pfffU4ssVw4/uAPf++fv3H4t7593B984HdhGrB1pR0NlETKuQQ20QNV0iSO61Lxv3qdD5SFuKb4tEJQW71izU/vuNckQAYRBIove7CtixXSAFWranQUocall/7glhtu0QND3lBN2JiJp0lEvwTzmEupsBKogKbQt2TSOlyyYNfh/Q/Y7+RTTjz66IP32vs9xsTgEtt8miUF6LOR6jN1m1BP+spsrEJENXzgj3/9pEVfXXTOqX++/z7VIQ1Ne0AhezJjUwqsdcZYoW3gYkmxuVXga0RfnZJw/JXRM04/9+WXGgl97ozFQXr5exxVV3zDCSP8UFjAZSlIlQfuf2rZ0uVA4Kzi7SmFK2CWZWGtFtYq6chGUx+JaF91tvuv7PH3n/vM+cvOOfZrh7xnr3foCioDIWcLPO+p4URl9OEmEj0f2yYQQM/rniuotTImDiuIUyQp/vCP3vsvx335W8f/y5998E8Gh6qobwyqQVjRcAahUmGYP1T0msPWEqesoY2q3nfHvReuvJibO45ixIecy8tPEVgLLoVkXqOenX76ma8891JYqdkGjRA8J8CaMelSyeg4lE5HR9MsHXzTbqioXd+48NP/8Lcnfv+Egw/+1Lt/7S1cQocXIIwwMvqqFApOen4iuk2NJor0Quw1etAL6m1PB3GJiQ2M0ogqqA4gzhCE+L3f/+1vHPf1M89ZdODffyqqRWmdy5WihbFJHbz8KmPBsBDmgM5PCOMQVC688JKbbrqDdkhrwGbecEw1lqSUzT33FStW3fbjn7ACQEGCvLIpgnB4iKtnODRE0lSr0Sc//bennnbSP332M+/6td25B4zprCvUY2SWHl6VDecyRT39kqX6RdFt6VkJI2cy2g+KseAhgwqgQ9QGZO9f/9Ujjzn4jLMXHfDJg97wpiHohD4yJPNjRkoVNdJQkBBQJjVBVKOt2vjKhtNPPWNspEEbB6sdPLXysg5oCrNM5n/H5pEHnzz3rPMtrZ3jpi+VYDNiWcCKzSVL01Gn0je+dZdP/cPfnLvk3COO+sw73rUHZ4JvX0OUX3aLkFbS8kmvVd4srFehiDJTmrEe/1I9rt921OOu3dDYRBwMEZokPzzWWcs4n1QIq3jvb/7KoUd+7vgTv7HvAX8S1mIEMZQFL7pLLhAJtQuQcuyUsxClEce/eOKZZResHh+BgmIp0sYo3mIjqbMJz8Vg/FHVS8+bc848P92UhqoCpYIgcFlKGmgViCLDDMSCGwHUETSG3lT9609/5MRF//q5L37yV961wHCdDZA5r7AD+czCedzB2VAQwWnqCGG+ga+EnVIAxWf3/qdvFN0WlAIIRwatSziyQG4nDFzKM0drBxbI7/ze3kcfe+jqNUs/csC+u7/tzXAJVy0VaJKQLNTkhM20wMWxrg2mqblg8ZKbrrs9EFgDWkH6/UqrLEsYrwQMeAMXLF5+4zU36WggrWfIJK0nOgpFi0ljZ2NGuCGlQ/+b7yKyUxoAAA+FSURBVP/Ngw/97EWXLjvk8P+992/sHlT93z9b5/UkayHwl5A31kd8XxS8HWsPjWV34Iv7+/mtPNLbQVv73lZzx7VzzmktYShKqSQhYVCr1d72tl/+6tcO+9bxx33sEwcMLqjYxiZI4hSPOWO69kk8Hg0NmDim6aLpOevMc9ev9yNMbomzWZoMVKtipT6WcsVb+8i65ctXq6CqJMJYXB1YoMKqiVN/wFEhaxInya//1/9yyGFf/OrXjvnbT338zW9eEIRoNJDSjSrIBLCeHe9Zfzyh+kPNHdfSWh6dM/D2JghEaz+KUYQgwnt+7VePPOrgsxef/ld/9/E3v3U3HWQSWqgUSI0x0QC3AKk17j8ffPikk06hF+8cNC8GKsgyW62Ez60bOf747zTGY2uRjtejN7yxMTouIgiEixePDPb6rb0POeKQb//bcZ/81Eff9Z638mRUaT7vdxiVigoCsE5D7dyOd2wWnpiNKtVsVNoLdVaroYgY40iFLHNp2hRjsWBX1IZJrz2O+crnTj7tux89cJ8Fuw4gGxvabaHJ6smrr0bDwyI6Ghy+4fqbf3zr7Uk9qYZQ0BvWj1fDahLjkjWX//TOn3lXPk1RCZNkHCG9oRTZ+N6/tdfnv/iP3/neNz/9v/bb8x27SOAXvkbsPAWtS5Ks0UioFbkVaiUlsXqBKzukQxxnxWoYRSqKJAhIM2ENTrlNY4n3mgMkxr57rz0OO/IL5yw+db+DPqoimiYbvmFB0hh3SSMZGVv/wkunnHz66Mb6yAYEGoMDA9wM3nXng8uWrRClkKV6MKLJAuKgYt/5rt2/8s2vnfCd//fXn9h/j7cvzDLWD1EIQ+gAVvyL7TDSAY/TlNfEWm9TMUevOWuxtNZKKa44HGCaFRoMEeiANsLwjYoTY8VVBpQOURnAO9+z+1e/fvgZZ53y8U99fM93vA3aohLooQEo9eiDjy4+e5lyiMfohoFbxe+feMqm9RtpABHCxCPhsH7f+9/7+S/9w7lLFu37kT8lpaqDyOhsE1rNfZ+lgbJiIbYwT+RTmmY0XYDisog5erH3c7NnIlJ0jHaLUaVIEhjngiDIuDSSccovTKmlIQEUggr22vvthx/+2X/99jc+e/Bna4MV0xghPXjv0ksuv/XWn9FLq49j5coLH374YdDd59Kn3Z7v3vNrXz/qX7/5tb/59AG7vbFCXyql4+7A6tloLiLKKaVEhA4cm9YqjKIgDAPAV8NwTsqcJRb5ROGY5WPqR5pGC3StM6c13S/tnKhAoPyRhBWGljtAHeE9e73181846LIfXPz5Iz6/57vfDu0aYyOLFp384IPP3HzzXcuXL7VJPRqu/sZvv+87J/7bmsuW7PfRP9nj7bsFERe+1HLFE+4pHaulWDgngGhjAdZOggn5RIY1j9xcYcQwB685S6wpx4orGuCHOg8BWioGXKeAIFD5oCPNkBr80ltrB318/++e+K0vHPJPf/RnH3jLL+321NOPvfDis+989zs++JH9jjrqiONP+Oaf/tnvRBEaCXeQgIAccnTHBdu+eI+y7ftz6I6aQ32ZblfEIReOMYVkU+REvk5Zmi4aLdGAxsI3BHyR98lPH/TPxx3zreO//ucf/MP9PvoXJ37/377+z8cecNBfvG3PBTpEkiHQ3BzA1+A0T79AikHEbSl5PpqXowFrSTNrs685kJh3xFLFoOZDJ4xD4AXKe/qOnjUXUGvB44lGg0cDaRhheEF11zcMirJhpIYXhvTG4sRt3JTVG1DaSyMBd4tkmNbkk2ct5v2l5iEC7LOnV97znFuMcTkUFWjNN4fCrZzhqUS1FgwvCFUgxmUpD06ramAwTAwMEFRkYDgIq5AAVpoOk1KwvAc6V01hvVOLwJO5CDE3LzU3u7XtXpFJhZBbFBZkkiGXQoYiUEoUPyJZZsfHM4Y6qjgnGzbFo/XMWke/isWyzDUaSZoarVGpgqaLO0XnPKVQXsC8I5Y3FVsNPLlViTRXwDg2Kfd2xEUJlIIKLII0FR2o2kBlaDhQWhqxowETDR0qcMPHTZ6jv2+cIGr/y2f4i8Rti0/Pp4+aa531btP2+iQTBTjwRYLGBg5a6zBUIiSL5RUEEAXhpTE2ltbr0NonadLSNGVhCk2dc7RbvEHPzK+MRYXTCKej6zSq6ckiqie12kGl6CDxiUl0gWO/mKbwxiSRgj+gdfG+Ed2jVhGePikBH3OGp09kjxJRXN1AtimkCQZqZJz3orSCOAlU6Etm9PlDYXOWmZC8tnbNkyOY4spLT5E/F7JUf3dC3A7onxduD/YOPAhMVZi0mCxTFZnHeX1OrHk8cj3e9ZJYPT5A/apeSax+Hbke17skVo8PUL+qVxKrX0eux/UuidXjA9Sv6vUrsfoV73mjd0mseTPUne1oSazO4j1vWiuJNW+GurMdLYnVWbznTWslsebNUHe2oyWxOov3vGltxoi1E4i1/y1U61kLUIpUodj2QidF6XZot8xo3wFYmDIpo6eiU6FRKKjgiAMlTxb/RiiPtgPbztys+5I/2C7V0UhL3VlrlCNNKaoXl/+jJSfiZHMcHUgpjw6JpfwfeFoWUGp7ITziaIaAbcHqxDpxjsACrLponXcpRbxnQgcxkCxHgwQq9LLMAWye4ABpIAQY5hk+4K0JmWBV8xF4QBxYpyIEvnwXPmr223Ts6Gu0Is171pOuGd+Rr62wc56gRQ1sOo/wu5A81WOB3Uof5tgWGi14ikK+a7aIbhluXnDLux1Pd4BY7DFlcs8mDTKjzTucryxGfSynoFVuOiEItLSqaH0365v8xYoLmZzZK3F2mbLZ7PNpdseRQ/4rv2eADMwpMprKs2AhRZrxPOJhAc2zbYOTZ3cyaKnSuTYJzERjvuN+OeOwk1iTlSGmLDadkMW2kEn1sOItbvZWUsDuU/w/im5qNqE9+QEi0BImKc1S7a+i+GbotYy48xNvwhdoP9KJSKHWLLbETk4S7/f4xgjQ5CEnsiyEAMRX4OHwBew0Ig7bvVjhdst0s0CuHxHYTAf2yzJDwU6Iy90mEKOW5DkgdBSWnhA+Do+e5yW6cs06sbbqlc07vFU2sSI6lK3vzIsc2eFeei4qzypC54U1MFkMKOM7XN90H5heuUKP6ZXd2VIWfuIwnEYFOTSEbLqyLQSn7BfnMWUaWnSwCKcShf7QRJvsO9ivvAuMF+KpwxxKXpCZzOGTDPMM+EjrbjOnm19dVsXj0+x+QbwibGbNq682t9oR7355CDhGk+abz9niwwKtHP8HaN5v6zqOk3Rq6Tbj32yjkKLmwmgYYzwIzOKZE4yfb2JE8bhGNPQ0RYSnVU5p/+sJNv+5M+fATNYKSC7o8SsMkf/tq++Cyf+kUWvxtFACJeyb44mel8BJkCe1Yw9zcb4AUxR2Fh42pZzzSWvZf+1/lwvduVTHm2WLyhEzrRVppMFjJ78rttxLJ86kBsa56YuzNiNH2YsgAIUR/3MLxJWxfpAsQ6E/GRZWIAJjU+v/FtaAOMA4WMcL/DjHZJ5JlAxSh8z5uyzDwilMiiw1zvIitqIRslJIV2DgMM92uw6gFK2wk01JU5MY73t5WJWBZDxelhCijOOsm55A2SgKSFGtOTyc+gTVKKXCsNUvtkwpGu/JkJoHQUAqxDEDQIGECCINP/uYdK2IhbZ5nDnwcZ+kpU+hM6gMIUAYYRXNmKPRQpZQXLc6rbrQcN7ZSkXTwIggs1mlEiJwkNS5hnOciLHDtAQuMTaNkzptPqtSCqyWxCp+2KMLXdvxJpPEkFicCSK0ScgMkiQ2JoZL4RJYhjHQAJiTi48k/hbvOt5qZdoGJCG9KtUwtYZ8pUT+v2zdcZ1m4olZJ5aQMJuLV9t5A5MklkZraOHA73/gdz/00b/88P77fOhjf/nhA/fZATlgv7/44P/YY4+3klK5sDVw4lN8K/3w4aLtnFdU/OVX83e+a8/9Pvyh/Q7Yd78D99vvwH32O5CRD+3rI/vs+1dMMrMQD9R+B+7z4fzWPvt/cN/99933Ix98//v/W20gSlM0GinXWV91Nz6zTawCM9vsGlOUPMH5RNMSRfjVd7zlf/7NQUceediXjz3ymC8ffuSxh+Ry6PTCwz7zj3/3znfuzrFxcJlJk4Rmz1UqPGvNm9kskM1SvZGgibXWJQnNlSPJOD3e9773fv7gfzr6mCOOPubQo4855Kgvf+moY790tA+/mCcPPeqYw476Mm8xPOxoHz+E+cd+5Yhjv3L4YUd+6WP7f/gtb4noXNFD4JrQrV7ONrHoM7Qcgxaliq5y8TLWNhoUt3Dh4C67VmtDqA0F1UFVHVLVQZlmuMsbhpIEcZxoLVprkpUzn9sr38rmLfqcLnmyedNTB9YA3PgpCUMNjUZqnWDX3QZrQ6o2JLUhVL0QDUobGRCfmi/AMk2pDhI97LJLVKvpsTHEOVObOEzd8uzmqtmtHoWRmKIV2pgwVEGgOFWV8isj7bbQ0PAUhqKINtz2QirPWc6RqFQirZHEGZ8JQ83KeasvhKpSc6rK7jMMQ04NFHEmLZErIOSuUIwT10RUYIt8n+NYkuKfEgQheGpTrQrNtiU6vNENoeKz26z1b3CEKFA8KFIEPgRXLwdamYwRcoiSQ1RgNZ2QqnvyKY9yg55rGIHNZWg+y0Qu6OFLqFvRayF36CDyTA/sFOGi8GZL6Nwr4Rxi2j/DL99rC/EiPk6DR/NHMEW4QfZVCSnmC3bhM+vE6kKf5myTLUL1QwdLYvXDKPWhjiWxZmHQyirhN20lDCUCM49AabFmHtOyRiJQEosglDLzCJTEmnlMyxqJQEksglDKzCNQEmvmMS1rJALzg1jsaCmdRaAkVmfxnjetlcSaN0Pd2Y6WxOos3vOmtZJY82aoO9vRklidxXvetFYSa94MdWc72iVidbaTZWudR6AkVucxnxctlsSaF8Pc+U6WxOo85vOixZJY82KYO9/Jklidx3xetFgSa14Mc+c72SJW51suW5zTCJTEmtPD273OlcTqHvZzuuWSWHN6eLvXuZJY3cN+TrdcEmtOD2/3OlcSq3vYd6flDrVaEqtDQM+3ZpQTlFIisNMIaJomBxEoRuAv0f5nvJRz/ve+yrBE4PUjYO0El5QWlFIisNMIBAKKhg+Vhiin8t8wUuL8VYYlAjuBAB8x1lpnaaus81+umfYro+OaqFCGJQI7jEBOGzCkffIhWaTAOB0vRbLlLLNl2CUE+hh5kgfNy/JbQB/LADReTpFfpZQI7DQC4n+v2Rahcla54mf9rVqwICylRGCnERgeDoaHw1yKSDTsc/T/BwAA//8f1ZmCAAAABklEQVQDAAQSR+FOl4ALAAAAAElFTkSuQmCC";

// A transparent white silhouette of the same real logo's "M" mark,
// used as a low-opacity decorative watermark on hero panels.
const MERCER_M_MASK_DATA_URL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAIGElEQVR4nO3dW6hc1R3H8V+C8W41itbG2mpKjDdsAl5qpFRpS6OkKKIgTUEftHlQRFB8KBZRn2xAlPqQB5X0IsUnhdJrVEqxCWrVgK1avNQL2koFgxfiJTo+rLP1ZDLzP3tm1tr//d/7+wExnpnZ++/MfGetnMucRYPBQABGW+w9ANBmBAIYCAQwEAhgIBDAQCCAgUAAA4EABgIBDAQCGAgEMBAIYCAQwEAggIFAAAOBAAYCAQwEAhgIBDAQCGAgEMBAIICBQAADgQAGAgEMBAIYCAQwEAhgIBDAQCCAgUAAA4EABgIBDAQCGAgEMBAIYCAQwEAggIFAAAOBAAYCAQyRAznIewB0X7RA9pY0mPvnHUkbfMdBDVv1xWO203mWiS0aDAbeM0ximaTXhz62VNKO5kdBDesk/W7oY4s8BplWtEC+KWn7iI+HutN7ZNSTK9RjFW2LNbx6VDY2OgXq+LP3ADlEW0GOlfTSmMtCvTJ13KmSHh9zWajHKdoKssu4LFTpHTcujnCiBbKQTd4DQB94D5BT1wLZIOkQ7yF67GJJ+3gPkVO0v4McLenVGtcLtc/tkDpPplCPTddWkMpvvAfooXBfBKyjq4Gsl3SE9xA9skHSvt5DlNDVLVYl1HIe1EFK3/ZTV6jHpKsrSOVm7wF64HnvAUrqeiA/k7TKe4gOu1HSl72HKKnrW6xKqGU9iCWSPpridqEei66vIJVbvQfooGe9B2hCXwK5XtJK7yE6ZKOkb3gP0YS+bLEqoZb3llos6ZMZbh/qMejLClJ5wHuADrC+YbRz+hbI+ZLO9B4isKsVbAWYVd+2WJVePciZLJf0YobjhLrv+7aCVH7vPUBAOeIIp6+BnCfpK95DBHKh9wBe+rrFqoRa7p3sL+n9jMcLdZ/3dQWp3O89QABvew/gqe+BXCDpRO8hWuw6pTfr662+b7EqoZb9hkz6bex1hbqv+76CVP7gPUALjXsPsl4hkORcSWd7D9EiPxFvDi6JLdawUMt/QSWfFKHuY1aQ3T3oPUALhHrFLI1Advdd9ft7tW7yHqBt2GKNFmobkMnBaubXSIS6b1lBRhv+nRZ9sMN7gDYikNHWSVrrPUSD7vEeoK3YYtlCbQemdLya/fnyUPcpK4jtXu8BGtCLN1+YFoHYfiTph95DFNSHF4CZsMWqJ9S2oKbTJT3qcN5Q9yUrSD3PeQ9QgEcc4RBIPSuVfsNuV9zuPUAUbLEmE2p7MEbTn7UaFuo+ZAWZzMveA2TAZ60mQCCT+brSrziO6ufeA0TDFms6obYJc1ZJesp7CAW771hBpvOk9wBTaEMc4RDIdFZL+r73EBPY4j1AVGyxZhNhu3CCpGe8h5gnwn32OVaQ2bTpiTdOhBlbi0Bmc4Kkn3oPYeDdWmbEFiuPNm4bDpD0nvcQI7TxvhqLFSSPf3sPMEIb4wiHQPI4TtJF3kPM81vvAbqCLVZebdg+rFK7v+bRhvuoNlaQvHZ4D6B2xxEOgeR1sKRLHM//a8dzdxJbrDIWq/l3KFwj6e8Nn3MabLGgnQ7njBBHOARSxj6SLm3wfPc1eK5eYYtV1uGS3ip8jrWS/lj4HDmF2mIRSFk7JC0tfI5QD6CCBcIWq6xDlH4PYim8r1VhrCDN2EvSJ5mPeY6khzMfswmsIAH9v/Dx3yxwzIhxhEMgyXqlV7YrCh3/MOX9rNbWjMcadpbSfbGp4DnCIJDkg7l/36X05PhxgXNsznScbyn/b8F6Q9Iypf/3Kr7Sq2oIBJIcMPTf9yo9WS7PfJ4cf3/aluEYlbeV3kjuKEn/HbpsRcbzhEUgyZIxH79bKZTVmc5ztKTrZ7h9rjd92ynpWEmHavzPsvwv07lCI5Bk7wUu364UyhkZznXrlLc7XenVfhZvKUWxvxZ+l0ieG+JOqNT93qnHlEL5wYzne22K28zybuy7lOI6XGlbhZoIZDp/UQrlyClv/1VJN0xw/eG/H9T1uNKcSzT5jwWH+gJZKQQymzeVnoBrprjtLTWvd5ImD/Gvko5R2pZhBgSSxzalUE6c8HZ1XqX/OcHxHpL0JaWvsr8y4SwYgUDyelaTh3KlcVndOP6m9C3235P07gTnxgIIpIwqlOU1rnvnmI9foLS9svxSKYzvSPqo7nCoj0DK+o9SKKcscL1RW637jes/Mnfcy0QYRRFIM55WekKvkPSvMde5at6fHxtznTskHSjp2/lGg4VAmvWCpJOVYnli6LJfKL2n1TWSThu67Nq521wj6f2SA2J3BOLnVElfk/SPeR+7WOkzUJXNSmHc1txYmG8v7wF67jXtuVqgRVhBAAOBAAYCAQwEAhgIBDAQSBLqrWjQHAIBDAQCGAgEMBAIYCAQwEAggIFAAAOBAAYCAQwEAhgIBDAQCGAgEMBAIICBQAADgSS81f+e+BkZxQuEJ3JzCETxAgEaRSAY51PvAdqAQAADgQAGAgEMBIJx+CyWCKTysfcALfSh9wBtEC2QUq9q+xY6bmT7eQ/QBtECKfVE5tVyT7u8B2iDaIGUelU7rNBxI1vqPUAbRAvkJUkvZz7mp5L+lPmYXfCrAsfcUuCYRS0aDPj2JmCcaCsI0CgCAQwEAhgIBDAQCGAgEMBAIICBQAADgQAGAgEMBAIYCAQwEAhgIBDAQCCAgUAAA4EABgIBDAQCGAgEMBAIYCAQwEAggIFAAAOBAAYCAQwEAhgIBDAQCGAgEMBAIICBQAADgQAGAgEMBAIYCAQwEAhgIBDAQCCAgUAAA4EABgIBDJ8Bd/MN8Ax2YmkAAAAASUVORK5CYII=";

function MercerLogoImg({ size = 36, radius = 8 }) {
  return (
    <img
      src={MERCER_LOGO_DATA_URL}
      alt="Mercer"
      width={size}
      height={size}
      style={{ borderRadius: radius, display: "block", objectFit: "cover" }}
    />
  );
}

// Shows a project's uploaded client logo (e.g. the client organization's own
// branding) if one has been set, falling back to Mercer's own logo — used
// wherever a survey/meeting's branding should reflect who it was built for.
function ClientLogo({ logoUrl, size = 36, radius = 8 }) {
  if (logoUrl) {
    return (
      <img
        src={logoUrl}
        alt="Client logo"
        width={size}
        height={size}
        style={{ borderRadius: radius, display: "block", objectFit: "contain", backgroundColor: "white" }}
      />
    );
  }
  return <MercerLogoImg size={size} radius={radius} />;
}

// ---------------------------------------------------------------------------
// Name / email helpers — initials and email are always derived from the
// name, never stored separately, so they can never fall out of sync.
// ---------------------------------------------------------------------------

function initialsFromName(name) {
  const parts = (name || "").trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

// mailto: links can silently fail inside sandboxed contexts (e.g. a preview
// iframe) — clicking one tries to navigate the frame itself to an
// unsupported protocol instead of launching a mail client, landing on a
// blank page. Copy-to-clipboard is a reliable alternative that works the
// same everywhere, and arguably a better UX for "get a colleague's email"
// than trying to force-launch whatever mail app happens to be configured.
function CopyEmailButton({ email, className }) {
  const [copied, setCopied] = useState(false);
  const handleCopy = async (e) => {
    e.preventDefault();
    try {
      await navigator.clipboard.writeText(email);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1500);
    } catch (err) {
      // clipboard API can fail without permission — nothing more to do
    }
  };
  return (
    <button type="button" onClick={handleCopy} className={className} title="Copy email address">
      {copied ? <CheckCircle2 size={12} /> : <Copy size={12} />} {copied ? "Copied!" : email}
    </button>
  );
}

// Audience for a meeting/survey can now be several cohorts at once (e.g.
// "Batch 1" and "Batch 2" invited to the same meeting), stored as an array.
// Still accepts a plain string for backward compatibility with anything
// created before this existed.
function audienceMatches(audience, cohort) {
  if (!audience) return true;
  const arr = Array.isArray(audience) ? audience : [audience];
  if (arr.length === 0) return true;
  return arr.includes("All Participants") || arr.includes(cohort);
}

function formatAudience(audience) {
  if (!audience) return "All Participants";
  const arr = Array.isArray(audience) ? audience : [audience];
  if (arr.length === 0 || arr.includes("All Participants")) return "All Participants";
  return arr.join(", ");
}

// A checkbox-based audience picker so a meeting or survey can target several
// cohorts at once instead of being limited to one (or everyone).
function AudienceSelector({ value, onChange, cohortOptions }) {
  const audienceValue = Array.isArray(value) ? value : value ? [value] : ["All Participants"];
  const isAll = audienceValue.includes("All Participants");
  const specificCohorts = cohortOptions.filter((c) => c !== "All Participants" && c !== "All Cohorts");

  const toggleAll = () => onChange(isAll ? [] : ["All Participants"]);
  const toggleCohort = (cohort) => {
    if (isAll) {
      onChange([cohort]);
    } else if (audienceValue.includes(cohort)) {
      onChange(audienceValue.filter((c) => c !== cohort));
    } else {
      onChange([...audienceValue, cohort]);
    }
  };

  return (
    <div className="space-y-1 rounded-lg border border-slate-200 p-2" style={{ maxHeight: 160, overflowY: "auto" }}>
      <label className="flex items-center gap-2 rounded px-2 py-1.5 text-sm hover:bg-slate-50">
        <input type="checkbox" checked={isAll} onChange={toggleAll} />
        <span className="font-medium text-slate-700">All Participants</span>
      </label>
      {specificCohorts.map((c) => (
        <label key={c} className={`flex items-center gap-2 rounded px-2 py-1.5 text-sm ${isAll ? "" : "hover:bg-slate-50"}`}>
          <input type="checkbox" checked={!isAll && audienceValue.includes(c)} disabled={isAll} onChange={() => toggleCohort(c)} />
          <span className={isAll ? "text-slate-300" : "text-slate-600"}>{c}</span>
        </label>
      ))}
      {specificCohorts.length === 0 && <p className="px-2 py-1 text-xs text-slate-400">No specific cohorts yet — add participants with a cohort set first.</p>}
    </div>
  );
}

function slugifyNamePart(s) {
  return s.toLowerCase().replace(/[^a-z]/g, "");
}

function domainFromEmail(email) {
  if (!email || !email.includes("@")) return "acmecorp.com";
  return email.split("@")[1] || "acmecorp.com";
}

function deriveEmail(name, domain) {
  const parts = (name || "").trim().split(/\s+/).filter(Boolean).map(slugifyNamePart).filter(Boolean);
  if (parts.length === 0) return `participant@${domain}`;
  if (parts.length === 1) return `${parts[0]}@${domain}`;
  return `${parts[0]}.${parts[parts.length - 1]}@${domain}`;
}

// ---------------------------------------------------------------------------
// Live app data — participants are React state (not constants), so editing
// a name, cohort, or any other field in the UI actually updates the app.
// ---------------------------------------------------------------------------

const AppDataContext = createContext(null);

function useAppData() {
  const ctx = useContext(AppDataContext);
  if (!ctx) throw new Error("useAppData must be used inside AppDataProvider");
  return ctx;
}

const PARTICIPANTS_STORAGE_KEY = "mercerassess-participants";
const SESSION_STORAGE_KEY = "mercerassess-session";
const ADMIN_ACCOUNTS_STORAGE_KEY = "mercerassess-admin-accounts";
const ADMIN_SESSION_STORAGE_KEY = "mercerassess-admin-session";
// Same demo password ("demo1234") and salt as the participant demo account —
// hashing only depends on password + salt, not what kind of account it is.
const adminAccountsSeed = [
  { id: "admin-1", name: "Admin Grader", email: "admin.grader@mercer.com", passwordSalt: "demoSalt123", passwordHash: "10a71203c4279d0fd55eacf9642862f1e491bf977af5b55653c9f42971c9f3a3" },
];
const PROJECTS_STORAGE_KEY = "mercerassess-projects";
const CURRENT_PROJECT_STORAGE_KEY = "mercerassess-current-project";
const projectsSeed = [{
  id: "default",
  name: "Default Project",
  criteria: [
    { id: "1", name: "Innovation", weight: 25, isReflectionCriterion: false },
    { id: "2", name: "Critical Thinking", weight: 20, isReflectionCriterion: false },
    { id: "3", name: "Reflection Assignment", weight: 20, isReflectionCriterion: true },
    { id: "4", name: "Depth of Analysis", weight: 25, isReflectionCriterion: false },
    { id: "5", name: "Grammar", weight: 10, isReflectionCriterion: false },
  ],
  groups: [],
  resources: [],
}];

// Computes a participant's overall score as the weighted average of their
// criteriaScores against a project's configured criteria weights — this is
// the real total (e.g. Assessment 50% + Attendance 20% + ALPs + Modules),
// not just the AI-graded reflection alone. Only numeric criteria scores
// count toward the calculation; qualitative values (like "Y"/"not due")
// are ignored here but still shown in the breakdown display. Normalizes by
// the weight actually used, so a participant missing a few criteria scores
// isn't unfairly penalized versus one with a full breakdown. Returns null
// if nothing numeric is available to compute from, so callers can fall
// back to a stored avgScore instead.
function computeWeightedTotal(participant, project) {
  const criteria = (project && project.criteria) || [];
  let weightedSum = 0;
  let weightUsed = 0;
  criteria.forEach((c) => {
    const raw = participant.criteriaScores ? participant.criteriaScores[c.name] : undefined;
    const num = typeof raw === "number" ? raw : parseFloat(raw);
    if (Number.isFinite(num)) {
      weightedSum += num * c.weight;
      weightUsed += c.weight;
    }
  });
  if (weightUsed === 0) return null;
  return Math.round((weightedSum / weightUsed) * 10) / 10;
}

// Generates and downloads a branded certificate as a PNG image, drawn on an
// in-memory canvas — this needs no external library (Canvas is built into
// every browser), so it works identically everywhere, and produces an actual
// image file rather than a PDF that might open in an unexpected app. Loads
// the real Mercer logo (used as a header mark and faint corner watermarks)
// before drawing, so this is async — resolves once the download has fired.
function generateCertificateImage({ participantName, programName, date, certId }) {
  return new Promise((resolve) => {
    const width = 1400;
    const height = 990;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");

    const finishDownload = () => {
      const fileSafeName = participantName.replace(/[^a-z0-9]+/gi, "_");
      canvas.toBlob((blob) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `Certificate_${fileSafeName}.png`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        resolve();
      }, "image/png");
    };

    const draw = (logoImg) => {
      // Background gradient, matching the app's brand gradient
      const gradient = ctx.createLinearGradient(0, 0, width, height);
      gradient.addColorStop(0, "#0b2559");
      gradient.addColorStop(0.55, "#0e3f7c");
      gradient.addColorStop(1, "#0f9a8e");
      ctx.fillStyle = gradient;
      ctx.fillRect(0, 0, width, height);

      // Faint large watermark logo, centered — repeated Mercer branding
      if (logoImg) {
        ctx.save();
        ctx.globalAlpha = 0.07;
        const wmSize = 620;
        ctx.drawImage(logoImg, width / 2 - wmSize / 2, height / 2 - wmSize / 2, wmSize, wmSize);
        ctx.restore();

        // Smaller corner marks, top-left and bottom-right
        ctx.save();
        ctx.globalAlpha = 0.14;
        ctx.drawImage(logoImg, 60, height - 150, 90, 90);
        ctx.drawImage(logoImg, width - 150, 60, 90, 90);
        ctx.restore();
      }

      // Decorative border
      ctx.strokeStyle = "rgba(255,255,255,0.85)";
      ctx.lineWidth = 3;
      ctx.strokeRect(46, 46, width - 92, height - 92);

      // Teal accent stripe at the bottom
      ctx.fillStyle = "#0f9a8e";
      ctx.fillRect(0, height - 24, width, 24);

      ctx.textAlign = "center";

      // Header logo mark (crisp, small, above the MERCER wordmark)
      if (logoImg) {
        const headerLogoSize = 64;
        ctx.drawImage(logoImg, width / 2 - headerLogoSize / 2, 80, headerLogoSize, headerLogoSize);
      }

      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 34px Helvetica, Arial, sans-serif";
      ctx.fillText("MERCER", width / 2, 195);

      ctx.fillStyle = "#c8d7eb";
      ctx.font = "20px Helvetica, Arial, sans-serif";
      ctx.fillText("CERTIFICATE OF COMPLETION", width / 2, 232);

      ctx.fillStyle = "#d2dcf0";
      ctx.font = "22px Helvetica, Arial, sans-serif";
      ctx.fillText("This certifies that", width / 2, 340);

      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 64px Georgia, 'Times New Roman', serif";
      ctx.fillText(participantName, width / 2, 440);

      ctx.fillStyle = "#d2dcf0";
      ctx.font = "22px Helvetica, Arial, sans-serif";
      ctx.fillText("has successfully completed", width / 2, 510);

      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 36px Helvetica, Arial, sans-serif";
      ctx.fillText(programName, width / 2, 570);

      ctx.textAlign = "left";
      ctx.fillStyle = "#b4c3dc";
      ctx.font = "16px Helvetica, Arial, sans-serif";
      ctx.fillText(`Issued ${date}`, 80, height - 65);

      ctx.textAlign = "right";
      ctx.fillText(`Certificate ID: ${certId}`, width - 80, height - 65);

      finishDownload();
    };

    const logoImg = new Image();
    logoImg.onload = () => draw(logoImg);
    logoImg.onerror = () => draw(null);
    logoImg.src = MERCER_LOGO_DATA_URL;
  });
}

// Opens LinkedIn's official "Add to Profile" flow for a certification, pre-
// filled with this program's details — this adds it to the participant's own
// Licenses & Certifications section (the standard, well-supported way to
// surface a credential on LinkedIn), rather than just sharing an image.
function openLinkedInAddToProfile({ programName, score, certId, issueDate }) {
  const [year, month] = issueDate.split("-");
  const params = new URLSearchParams({
    startTask: "CERTIFICATION_NAME",
    name: `${programName} — Mercer`,
    organizationName: "Mercer",
    issueYear: year,
    issueMonth: String(Number(month)),
    certId,
  });
  window.open(`https://www.linkedin.com/profile/add?${params.toString()}`, "_blank", "noopener,noreferrer");
}

// Password hashing via the browser's built-in Web Crypto API — no library
// needed, and it's a standard browser API so this works identically whether
// deployed on Vercel or run here in the Claude preview. This is a real
// (if minimal) improvement over storing plaintext passwords, but this is
// still a demo app: there's no email verification, no rate limiting on
// login attempts, and no server-side session/cookie security. Treat this as
// "better than plaintext," not as production-grade authentication.
function toHex(buffer) {
  return Array.from(new Uint8Array(buffer)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

function randomSalt() {
  return toHex(window.crypto.getRandomValues(new Uint8Array(16)));
}

async function hashPassword(password, salt) {
  const data = new TextEncoder().encode(salt + password);
  const digest = await window.crypto.subtle.digest("SHA-256", data);
  return toHex(digest);
}

function AppDataProvider({ children }) {
  const [participants, setParticipants] = useState(participantsSeed);
  const [toast, setToast] = useState(null);
  const [loaded, setLoaded] = useState(false); // becomes true once the initial storage read finishes
  const [sessionEmployeeId, setSessionEmployeeId] = useState(() => {
    try {
      return window.localStorage.getItem(SESSION_STORAGE_KEY) || window.sessionStorage.getItem(SESSION_STORAGE_KEY) || null;
    } catch (err) { return null; }
  });

  // Admin accounts — real, password-verified accounts (not the anyone-can-
  // click-through mock this used to be). Loaded/saved the same way as
  // participants; session tracked separately so an admin login and a
  // participant login on the same device don't clash.
  const [adminAccounts, setAdminAccounts] = useState(adminAccountsSeed);
  const [adminAccountsLoaded, setAdminAccountsLoaded] = useState(false);
  const [adminSessionId, setAdminSessionId] = useState(() => {
    try { return window.localStorage.getItem(ADMIN_SESSION_STORAGE_KEY) || null; } catch (err) { return null; }
  });

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const result = await storage.get(ADMIN_ACCOUNTS_STORAGE_KEY);
        if (!cancelled && result && result.value) {
          const parsed = JSON.parse(result.value);
          if (Array.isArray(parsed) && parsed.length > 0) setAdminAccounts(parsed);
        }
      } catch (err) {
        // nothing saved yet — keep the seed default admin account
      } finally {
        if (!cancelled) setAdminAccountsLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!adminAccountsLoaded) return;
    (async () => {
      try {
        await storage.set(ADMIN_ACCOUNTS_STORAGE_KEY, JSON.stringify(adminAccounts));
      } catch (err) {
        console.error("Failed to save admin accounts", err);
        showToast("Couldn't save — check the database connection (see README)");
      }
    })();
  }, [adminAccounts, adminAccountsLoaded]);

  const currentAdmin = adminAccounts.find((a) => a.id === adminSessionId) || null;

  const adminLogin = (id) => {
    try { window.localStorage.setItem(ADMIN_SESSION_STORAGE_KEY, id); } catch (err) { /* ignore */ }
    setAdminSessionId(id);
  };

  const adminLogout = () => {
    try { window.localStorage.removeItem(ADMIN_SESSION_STORAGE_KEY); } catch (err) { /* ignore */ }
    setAdminSessionId(null);
  };

  const fetchFreshAdminAccounts = async () => {
    try {
      const result = await storage.get(ADMIN_ACCOUNTS_STORAGE_KEY);
      if (result && result.value) {
        const parsed = JSON.parse(result.value);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (err) {
      // fall through
    }
    return adminAccounts;
  };

  const verifyAdminLogin = async (email, password) => {
    const normalized = (email || "").trim().toLowerCase();
    try {
      const freshList = await fetchFreshAdminAccounts();
      const match = freshList.find((a) => a.email.toLowerCase() === normalized);
      if (!match || !match.passwordHash) return { ok: false, error: "No admin account found for this email." };
      const hash = await hashPassword(password || "", match.passwordSalt);
      if (hash !== match.passwordHash) return { ok: false, error: "Incorrect password." };
      setAdminAccounts(freshList);
      adminLogin(match.id);
      return { ok: true };
    } catch (err) {
      console.error("Admin login failed:", err);
      return { ok: false, error: "Something went wrong signing in. Please try again in a moment." };
    }
  };

  // Only an already-signed-in admin can create another admin account —
  // there's no public "sign up as admin" path, unlike participants.
  const addAdminAccount = async (name, email, password) => {
    const trimmedEmail = (email || "").trim();
    const normalized = trimmedEmail.toLowerCase();
    if (!name || !name.trim()) return { ok: false, error: "Enter a name." };
    if (!trimmedEmail) return { ok: false, error: "Enter an email." };
    if (!password || password.length < 6) return { ok: false, error: "Password must be at least 6 characters." };
    try {
      const freshList = await fetchFreshAdminAccounts();
      if (freshList.some((a) => a.email.toLowerCase() === normalized)) {
        return { ok: false, error: "An admin account already exists for this email." };
      }
      const salt = randomSalt();
      const hash = await hashPassword(password, salt);
      const newAccount = { id: `admin-${Date.now()}`, name: name.trim(), email: trimmedEmail, passwordSalt: salt, passwordHash: hash };
      const nextList = [...freshList, newAccount];
      await storage.set(ADMIN_ACCOUNTS_STORAGE_KEY, JSON.stringify(nextList));
      setAdminAccounts(nextList);
      showToast(`Added admin account for ${name.trim()}`);
      return { ok: true };
    } catch (err) {
      console.error("Failed to add admin account:", err);
      return { ok: false, error: "Something went wrong. Please try again." };
    }
  };

  const removeAdminAccount = async (id) => {
    if (adminAccounts.length <= 1) {
      showToast("Can't remove the last admin account");
      return;
    }
    const nextList = adminAccounts.filter((a) => a.id !== id);
    try {
      await storage.set(ADMIN_ACCOUNTS_STORAGE_KEY, JSON.stringify(nextList));
      setAdminAccounts(nextList);
      showToast("Admin account removed");
    } catch (err) {
      console.error("Failed to remove admin account:", err);
      showToast("Couldn't save — check the database connection (see README)");
    }
  };

  const requestAdminPasswordReset = async (email) => {
    const normalized = (email || "").trim().toLowerCase();
    const generic = { ok: true, message: "If an admin account exists for that email, a reset link has been sent." };
    try {
      const freshList = await fetchFreshAdminAccounts();
      const match = freshList.find((a) => a.email.toLowerCase() === normalized);
      if (!match) return generic;

      const token = Array.from(crypto.getRandomValues(new Uint8Array(24))).map((b) => b.toString(16).padStart(2, "0")).join("");
      const expiry = Date.now() + 60 * 60 * 1000;
      const nextList = freshList.map((a) => (a.id === match.id ? { ...a, resetToken: token, resetTokenExpiry: expiry } : a));
      await storage.set(ADMIN_ACCOUNTS_STORAGE_KEY, JSON.stringify(nextList));
      setAdminAccounts(nextList);

      const resetUrl = `${window.location.origin}${window.location.pathname}?adminResetToken=${token}&adminResetEmail=${encodeURIComponent(match.email)}`;
      await fetch("/api/send-reminder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          recipients: [match.email],
          subject: "Reset your MercerAssess admin password",
          message: `Hi ${match.name},\n\nSomeone requested a password reset for your admin account. Click the link below to set a new password — this link expires in 1 hour.\n\n${resetUrl}\n\nIf you didn't request this, you can ignore this email.`,
        }),
      });
      return generic;
    } catch (err) {
      console.error("Admin password reset request failed:", err);
      return generic;
    }
  };

  const resetAdminPassword = async (email, token, newPassword) => {
    const normalized = (email || "").trim().toLowerCase();
    if (!newPassword || newPassword.length < 6) return { ok: false, error: "Password must be at least 6 characters." };
    try {
      const freshList = await fetchFreshAdminAccounts();
      const match = freshList.find((a) => a.email.toLowerCase() === normalized);
      if (!match || !match.resetToken || match.resetToken !== token) return { ok: false, error: "This reset link is invalid. Request a new one." };
      if (!match.resetTokenExpiry || Date.now() > match.resetTokenExpiry) return { ok: false, error: "This reset link has expired. Request a new one." };

      const salt = randomSalt();
      const hash = await hashPassword(newPassword, salt);
      const nextList = freshList.map((a) => (a.id === match.id ? { ...a, passwordSalt: salt, passwordHash: hash, resetToken: null, resetTokenExpiry: null } : a));
      await storage.set(ADMIN_ACCOUNTS_STORAGE_KEY, JSON.stringify(nextList));
      setAdminAccounts(nextList);
      adminLogin(match.id);
      return { ok: true };
    } catch (err) {
      console.error("Admin password reset failed:", err);
      return { ok: false, error: "Something went wrong. Please try again." };
    }
  };

  // Projects let separate concurrent programs (e.g. two client engagements
  // running at once) keep their own participant rosters instead of one
  // shared list. Which project is "current" is a per-device UI preference
  // (like the login session), not shared data, so it lives in localStorage.
  const [projects, setProjects] = useState(projectsSeed);
  const [projectsLoaded, setProjectsLoaded] = useState(false);
  const [currentProjectId, setCurrentProjectIdState] = useState(() => {
    try { return window.localStorage.getItem(CURRENT_PROJECT_STORAGE_KEY) || "default"; } catch (err) { return "default"; }
  });

  const setCurrentProjectId = (id) => {
    try { window.localStorage.setItem(CURRENT_PROJECT_STORAGE_KEY, id); } catch (err) { /* ignore */ }
    setCurrentProjectIdState(id);
  };

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const result = await storage.get(PROJECTS_STORAGE_KEY);
        if (!cancelled && result && result.value) {
          const parsed = JSON.parse(result.value);
          if (Array.isArray(parsed) && parsed.length > 0) setProjects(parsed);
        }
      } catch (err) {
        // nothing saved yet — keep the seed default project
      } finally {
        if (!cancelled) setProjectsLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!projectsLoaded) return;
    (async () => {
      try {
        await storage.set(PROJECTS_STORAGE_KEY, JSON.stringify(projects));
      } catch (err) {
        console.error("Failed to save projects", err);
        showToast("Couldn't save — check the database connection (see README)");
      }
    })();
  }, [projects, projectsLoaded]);

  const addProject = (name) => {
    const trimmed = (name || "").trim();
    if (!trimmed) return;
    const id = `proj-${Date.now()}`;
    setProjects((prev) => [
      ...prev,
      {
        id,
        name: trimmed,
        criteria: [
          { id: `${Date.now()}-1`, name: "Criterion 1", weight: 50, isReflectionCriterion: false },
          { id: `${Date.now()}-2`, name: "Criterion 2", weight: 50, isReflectionCriterion: false },
        ],
        groups: [],
        resources: [],
      },
    ]);
    setCurrentProjectId(id);
    showToast(`Created project "${trimmed}" — add participants and set up its grading criteria next`);
    return id;
  };

  const updateProject = (id, data) => {
    setProjects((prev) => prev.map((p) => (p.id === id ? { ...p, ...data } : p)));
  };

  // Deleting a project never silently discards its participants — they're
  // reassigned to Default Project so their data (including any imported
  // grades) is preserved, just no longer grouped under the deleted project.
  const deleteProject = (id) => {
    if (id === "default") return;
    setParticipants((prev) => prev.map((p) => ((p.projectId || "default") === id ? { ...p, projectId: "default" } : p)));
    setProjects((prev) => prev.filter((p) => p.id !== id));
    if (currentProjectId === id) setCurrentProjectId("default");
    showToast("Project deleted — its participants were moved to Default Project");
  };

  // Merges freshly-loaded stored data with the built-in seed defaults, so
  // fields added to the app in a later update (like passwordHash/passwordSalt
  // for the demo account) aren't silently missing just because someone's
  // data was saved before those fields existed. Matches by employeeId; any
  // field actually present in the stored record wins over the seed default.
  const mergeWithSeedDefaults = (list) => {
    const seedById = new Map(participantsSeed.map((p) => [p.employeeId, p]));
    return list.map((p) => {
      const seedMatch = seedById.get(p.employeeId);
      return seedMatch ? { ...seedMatch, ...p } : p;
    });
  };

  // Load any previously-saved edits once, on mount. If nothing was saved yet
  // (or the read fails), we just keep the built-in sample data.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const result = await storage.get(PARTICIPANTS_STORAGE_KEY);
        if (!cancelled && result && result.value) {
          const parsed = JSON.parse(result.value);
          if (Array.isArray(parsed) && parsed.length > 0) setParticipants(mergeWithSeedDefaults(parsed));
        }
      } catch (err) {
        // key doesn't exist yet on first-ever visit — that's expected, keep the seed data
      } finally {
        if (!cancelled) setLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  // Persist to storage every time participants change, but only after the
  // initial load has completed (otherwise we'd overwrite saved data with the
  // seed data for a split second on every mount).
  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try {
        await storage.set(PARTICIPANTS_STORAGE_KEY, JSON.stringify(participants));
      } catch (err) {
        console.error("Failed to save participant changes", err);
        showToast("Couldn't save — check the database connection (see README)");
      }
    })();
  }, [participants, loaded]);

  const showToast = (message) => {
    setToast(message);
    window.setTimeout(() => setToast((t) => (t === message ? null : t)), 2600);
  };

  // Login session is stored in real browser storage (not the shared KV
  // `storage` shim) since who's logged in on this device is inherently a
  // per-device concern, distinct from the shared participant data itself.
  // "Remember me" controls WHICH storage: localStorage persists across
  // browser restarts, sessionStorage clears when the tab/browser closes —
  // previously this checkbox was purely decorative and did neither.
  const login = (employeeId, remember = true) => {
    try {
      if (remember) {
        window.localStorage.setItem(SESSION_STORAGE_KEY, employeeId);
        window.sessionStorage.removeItem(SESSION_STORAGE_KEY);
      } else {
        window.sessionStorage.setItem(SESSION_STORAGE_KEY, employeeId);
        window.localStorage.removeItem(SESSION_STORAGE_KEY);
      }
    } catch (err) { /* ignore */ }
    setSessionEmployeeId(employeeId);
  };

  const logout = () => {
    try {
      window.localStorage.removeItem(SESSION_STORAGE_KEY);
      window.sessionStorage.removeItem(SESSION_STORAGE_KEY);
    } catch (err) { /* ignore */ }
    setSessionEmployeeId(null);
  };

  const currentUser = participants.find((p) => p.employeeId === sessionEmployeeId) || null;

  const updateParticipant = (employeeId, data) => {
    setParticipants((prev) => prev.map((p) => (p.employeeId === employeeId ? { ...p, ...data } : p)));
    showToast(`Saved changes for ${data.name || "participant"}`);
  };

  const addParticipant = (data) => {
    const employeeId = data.employeeId && data.employeeId.trim()
      ? data.employeeId.trim()
      : `EMP-${Math.floor(1000 + Math.random() * 9000)}`;
    setParticipants((prev) => [
      ...prev,
      { count: 0, last: new Date().toISOString().slice(0, 10), status: "Active", avgScore: 0, passwordHash: null, passwordSalt: null, criteriaScores: {}, projectId: currentProjectId, ...data, employeeId },
    ]);
    showToast(`Added ${data.name}`);
  };

  const deleteParticipant = (employeeId, name) => {
    setParticipants((prev) => prev.filter((p) => p.employeeId !== employeeId));
    if (employeeId === sessionEmployeeId) logout();
    showToast(`Removed ${name || "participant"}`);
  };

  const resetParticipants = async () => {
    setParticipants(participantsSeed);
    try {
      await storage.delete(PARTICIPANTS_STORAGE_KEY);
    } catch (err) {
      // nothing saved yet — fine
    }
    logout();
    showToast("Reset to sample data");
  };

  // Fetches the current participants list directly from storage, bypassing
  // the (possibly stale) in-memory React state. Used right before signup/
  // login checks so two people acting within moments of each other on
  // different devices are checked against the freshest data available,
  // narrowing (though not fully eliminating, without per-record atomic
  // writes) the window for one device's save to silently clobber another's.
  const fetchFreshParticipants = async () => {
    try {
      const result = await storage.get(PARTICIPANTS_STORAGE_KEY);
      if (result && result.value) {
        const parsed = JSON.parse(result.value);
        if (Array.isArray(parsed) && parsed.length > 0) return mergeWithSeedDefaults(parsed);
      }
    } catch (err) {
      // fall through to whatever's currently in memory
    }
    return participants;
  };

  // Self-service signup. If the email already belongs to a participant
  // record (e.g. one an admin added ahead of time) without a password yet,
  // this claims that record — setting the password and updating the name to
  // whatever they entered. If the email isn't registered at all, it creates
  // a brand-new participant record from scratch, so anyone can sign up
  // without needing an admin to add them first.
  const signup = async (name, email, password) => {
    const normalized = (email || "").trim().toLowerCase();
    if (!name || !name.trim()) return { ok: false, error: "Enter your name." };
    if (!password || password.length < 6) return { ok: false, error: "Password must be at least 6 characters." };

    try {
      const freshList = await fetchFreshParticipants();
      const existing = freshList.find((p) => p.email.toLowerCase() === normalized);
      if (existing && existing.passwordHash) return { ok: false, error: "An account already exists for this email. Try signing in instead." };

      const salt = randomSalt();
      const hash = await hashPassword(password, salt);

      if (existing) {
        const nextList = freshList.map((p) => (p.employeeId === existing.employeeId ? { ...p, name: name.trim(), passwordSalt: salt, passwordHash: hash } : p));
        await storage.set(PARTICIPANTS_STORAGE_KEY, JSON.stringify(nextList));
        setParticipants(nextList);
        login(existing.employeeId);
        showToast(`Account created for ${name.trim()}`);
        return { ok: true };
      }

      const verificationToken = Array.from(crypto.getRandomValues(new Uint8Array(24))).map((b) => b.toString(16).padStart(2, "0")).join("");
      const employeeId = `EMP-${Math.floor(1000 + Math.random() * 9000)}`;
      const newParticipant = {
        employeeId,
        name: name.trim(),
        email: (email || "").trim(),
        cohort: "Cohort A",
        dept: "",
        enrolled: new Date().toISOString().slice(0, 10),
        count: 0,
        last: new Date().toISOString().slice(0, 10),
        status: "Active",
        avgScore: 0,
        passwordSalt: salt,
        passwordHash: hash,
        emailVerified: false,
        verificationToken,
      };
      const nextList = [...freshList, newParticipant];
      await storage.set(PARTICIPANTS_STORAGE_KEY, JSON.stringify(nextList));
      setParticipants(nextList);
      login(employeeId);
      showToast(`Account created for ${name.trim()}`);

      // Verification email is sent best-effort — a failure here shouldn't
      // block account creation, since the account still works (just
      // unverified) and there's a "resend" option available afterward.
      try {
        const verifyUrl = `${window.location.origin}${window.location.pathname}?verifyToken=${verificationToken}&verifyEmail=${encodeURIComponent(newParticipant.email)}`;
        await fetch("/api/send-reminder", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            recipients: [newParticipant.email],
            subject: "Verify your MercerAssess email address",
            message: `Hi ${name.trim()},\n\nWelcome to MercerAssess! Please verify your email address by clicking the link below.\n\n${verifyUrl}\n\nUntil verified, some actions (like meeting check-in) will be unavailable.`,
          }),
        });
      } catch (err) {
        // non-fatal
      }

      return { ok: true };
    } catch (err) {
      console.error("Signup failed:", err);
      return { ok: false, error: "Couldn't save your account — the shared database may not be connected yet. Please try again in a moment." };
    }
  };

  const verifyEmail = async (email, token) => {
    const normalized = (email || "").trim().toLowerCase();
    try {
      const freshList = await fetchFreshParticipants();
      const match = freshList.find((p) => p.email.toLowerCase() === normalized);
      if (!match || !match.verificationToken || match.verificationToken !== token) {
        return { ok: false, error: "This verification link is invalid or has already been used." };
      }
      const nextList = freshList.map((p) => (p.employeeId === match.employeeId ? { ...p, emailVerified: true, verificationToken: null } : p));
      await storage.set(PARTICIPANTS_STORAGE_KEY, JSON.stringify(nextList));
      setParticipants(nextList);
      login(match.employeeId);
      return { ok: true };
    } catch (err) {
      console.error("Email verification failed:", err);
      return { ok: false, error: "Something went wrong. Please try again." };
    }
  };

  const resendVerificationEmail = async (employeeId) => {
    try {
      const freshList = await fetchFreshParticipants();
      const match = freshList.find((p) => p.employeeId === employeeId);
      if (!match || match.emailVerified) return { ok: false, error: "This account is already verified." };

      const token = Array.from(crypto.getRandomValues(new Uint8Array(24))).map((b) => b.toString(16).padStart(2, "0")).join("");
      const nextList = freshList.map((p) => (p.employeeId === employeeId ? { ...p, verificationToken: token } : p));
      await storage.set(PARTICIPANTS_STORAGE_KEY, JSON.stringify(nextList));
      setParticipants(nextList);

      const verifyUrl = `${window.location.origin}${window.location.pathname}?verifyToken=${token}&verifyEmail=${encodeURIComponent(match.email)}`;
      await fetch("/api/send-reminder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          recipients: [match.email],
          subject: "Verify your MercerAssess email address",
          message: `Hi ${match.name},\n\nHere's your verification link.\n\n${verifyUrl}`,
        }),
      });
      return { ok: true };
    } catch (err) {
      console.error("Failed to resend verification email:", err);
      return { ok: false, error: "Something went wrong. Please try again." };
    }
  };

  const verifyLogin = async (email, password, remember = true) => {
    const normalized = (email || "").trim().toLowerCase();
    try {
      const freshList = await fetchFreshParticipants();
      const match = freshList.find((p) => p.email.toLowerCase() === normalized);
      if (!match || !match.passwordHash) return { ok: false, error: "No account found for this email. Create one first." };
      const hash = await hashPassword(password || "", match.passwordSalt);
      if (hash !== match.passwordHash) return { ok: false, error: "Incorrect password." };
      setParticipants(freshList);
      login(match.employeeId, remember);
      return { ok: true };
    } catch (err) {
      console.error("Login failed:", err);
      return { ok: false, error: "Something went wrong signing in. Please try again in a moment." };
    }
  };

  // Sends a password reset link by email (reusing the same Resend-based
  // route meeting reminders use), valid for 1 hour. Always returns a
  // generic success message regardless of whether the email matched an
  // account, so this can't be used to test which emails have accounts.
  const requestPasswordReset = async (email) => {
    const normalized = (email || "").trim().toLowerCase();
    const generic = { ok: true, message: "If an account exists for that email, a reset link has been sent." };
    try {
      const freshList = await fetchFreshParticipants();
      const match = freshList.find((p) => p.email.toLowerCase() === normalized);
      if (!match) return generic;

      const token = Array.from(crypto.getRandomValues(new Uint8Array(24))).map((b) => b.toString(16).padStart(2, "0")).join("");
      const expiry = Date.now() + 60 * 60 * 1000; // 1 hour
      const nextList = freshList.map((p) => (p.employeeId === match.employeeId ? { ...p, resetToken: token, resetTokenExpiry: expiry } : p));
      await storage.set(PARTICIPANTS_STORAGE_KEY, JSON.stringify(nextList));
      setParticipants(nextList);

      const resetUrl = `${window.location.origin}${window.location.pathname}?resetToken=${token}&resetEmail=${encodeURIComponent(match.email)}`;
      await fetch("/api/send-reminder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          recipients: [match.email],
          subject: "Reset your MercerAssess password",
          message: `Hi ${match.name},\n\nSomeone requested a password reset for your account. Click the link below to set a new password — this link expires in 1 hour.\n\n${resetUrl}\n\nIf you didn't request this, you can ignore this email.`,
        }),
      });
      return generic;
    } catch (err) {
      console.error("Password reset request failed:", err);
      return generic; // still generic — don't reveal anything on failure either
    }
  };

  const resetPassword = async (email, token, newPassword) => {
    const normalized = (email || "").trim().toLowerCase();
    if (!newPassword || newPassword.length < 6) return { ok: false, error: "Password must be at least 6 characters." };
    try {
      const freshList = await fetchFreshParticipants();
      const match = freshList.find((p) => p.email.toLowerCase() === normalized);
      if (!match || !match.resetToken || match.resetToken !== token) return { ok: false, error: "This reset link is invalid. Request a new one." };
      if (!match.resetTokenExpiry || Date.now() > match.resetTokenExpiry) return { ok: false, error: "This reset link has expired. Request a new one." };

      const salt = randomSalt();
      const hash = await hashPassword(newPassword, salt);
      const nextList = freshList.map((p) => (p.employeeId === match.employeeId ? { ...p, passwordSalt: salt, passwordHash: hash, resetToken: null, resetTokenExpiry: null } : p));
      await storage.set(PARTICIPANTS_STORAGE_KEY, JSON.stringify(nextList));
      setParticipants(nextList);
      login(match.employeeId);
      return { ok: true };
    } catch (err) {
      console.error("Password reset failed:", err);
      return { ok: false, error: "Something went wrong. Please try again." };
    }
  };

  return (
    <AppDataContext.Provider value={{ participants, currentUser, updateParticipant, addParticipant, deleteParticipant, resetParticipants, showToast, login, logout, signup, verifyLogin, requestPasswordReset, resetPassword, verifyEmail, resendVerificationEmail, projects, currentProjectId, setCurrentProjectId, addProject, updateProject, deleteProject, adminAccounts, currentAdmin, adminLogout, verifyAdminLogin, addAdminAccount, removeAdminAccount, requestAdminPasswordReset, resetAdminPassword }}>
      {children}
      {toast && (
        <div className="fixed bottom-6 left-1/2 z-50 flex -translate-x-1/2 items-center gap-2 rounded-full bg-slate-900 px-4 py-2.5 text-sm font-medium text-white shadow-lg">
          <CheckCircle2 size={15} className="text-emerald-400" /> {toast}
        </div>
      )}
    </AppDataContext.Provider>
  );
}

// ---------------------------------------------------------------------------
// Reusable participant add/edit modal — the core of "make names editable".
// Changing the name box live-regenerates the email unless the person has
// typed into the email box themselves, in which case a small link lets them
// snap back to the auto-generated address.
// ---------------------------------------------------------------------------

// Native window.confirm() is unreliable inside this artifact's sandboxed
// iframe (it can silently no-op or throw), the same category of issue as the
// login form's native HTML5 validation earlier. This in-app modal replaces
// every confirm() call so destructive actions always actually prompt.
function ConfirmModal({ title = "Are you sure?", message, confirmLabel = "Delete", onConfirm, onCancel }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-sm rounded-2xl bg-white p-6 shadow-xl">
        <h3 className="text-lg font-bold text-slate-800">{title}</h3>
        <p className="mt-2 text-sm text-slate-500">{message}</p>
        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onCancel} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Cancel
          </button>
          <button onClick={onConfirm} className="rounded-lg bg-red-600 px-4 py-2 text-sm font-semibold text-white hover:bg-red-700">
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

function ParticipantFormModal({ initial, title, showStatus, onClose, onSave }) {
  const { updateParticipant } = useAppData();
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const [form, setForm] = useState(() => ({
    name: initial?.name || "",
    email: initial?.email || "",
    cohort: initial?.cohort || "Cohort A",
    dept: initial?.dept || "",
    employeeId: initial?.employeeId || "",
    enrolled: initial?.enrolled || new Date().toISOString().slice(0, 10),
    status: initial?.status || "Active",
  }));
  const [autoEmail, setAutoEmail] = useState(true);

  const handleNameChange = (value) => {
    setForm((f) => ({
      ...f,
      name: value,
      email: autoEmail ? deriveEmail(value, domainFromEmail(f.email)) : f.email,
    }));
  };

  const handleEmailChange = (value) => {
    setAutoEmail(false);
    setForm((f) => ({ ...f, email: value }));
  };

  const resetAutoEmail = () => {
    setAutoEmail(true);
    setForm((f) => ({ ...f, email: deriveEmail(f.name, domainFromEmail(f.email)) }));
  };

  const canSave = form.name.trim().length > 0 && form.email.trim().length > 0;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">{title}</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>

        <div className="mb-4 flex items-center gap-3">
          <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-full bg-blue-800 text-sm font-semibold text-white">
            {initialsFromName(form.name)}
          </div>
          <p className="text-xs text-slate-400">Initials and email update automatically as you type the name.</p>
        </div>

        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Full name</label>
            <input
              autoFocus
              value={form.name}
              onChange={(e) => handleNameChange(e.target.value)}
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
              placeholder="e.g. Mokhtar Alkhlifa"
            />
          </div>
          <div>
            <div className="mb-1 flex items-center justify-between">
              <label className="block text-xs font-semibold uppercase text-slate-400">Email</label>
              {!autoEmail && (
                <button onClick={resetAutoEmail} className="text-xs font-medium text-blue-700 hover:underline">
                  Use name-based email
                </button>
              )}
            </div>
            <input
              value={form.email}
              onChange={(e) => handleEmailChange(e.target.value)}
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
              placeholder="name@company.com"
            />
            {autoEmail && <p className="mt-1 text-xs text-slate-400">Generated automatically from the name above.</p>}
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Cohort</label>
              <select
                value={form.cohort}
                onChange={(e) => setForm((f) => ({ ...f, cohort: e.target.value }))}
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"
              >
                <option>Cohort A</option>
                <option>Cohort B</option>
                <option>Cohort C</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Employee ID</label>
              <input
                value={form.employeeId}
                onChange={(e) => setForm((f) => ({ ...f, employeeId: e.target.value }))}
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"
                placeholder="EMP-1234"
              />
            </div>
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Department / Track</label>
            <input
              value={form.dept}
              onChange={(e) => setForm((f) => ({ ...f, dept: e.target.value }))}
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"
              placeholder="e.g. Strategy & Innovation"
            />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Enrolled</label>
              <input
                type="date"
                value={form.enrolled}
                onChange={(e) => setForm((f) => ({ ...f, enrolled: e.target.value }))}
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"
              />
            </div>
            {showStatus && (
              <div>
                <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Status</label>
                <select
                  value={form.status}
                  onChange={(e) => setForm((f) => ({ ...f, status: e.target.value }))}
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm"
                >
                  <option>Active</option>
                  <option>Inactive</option>
                </select>
              </div>
            )}
          </div>
        </div>

        {initial && initial.criteriaScores && Object.keys(initial.criteriaScores).length > 0 && (
          <div className="mt-4 rounded-lg bg-slate-50 p-3">
            <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Grading Criteria (from last import)</div>
            <div className="grid grid-cols-2 gap-1.5" style={{ maxHeight: 160, overflowY: "auto" }}>
              {Object.entries(initial.criteriaScores).map(([label, value]) => (
                <div key={label} className="flex items-center justify-between gap-2 rounded bg-white px-2 py-1.5 text-xs">
                  <span className="truncate text-slate-500" title={label}>{label}</span>
                  <span className="shrink-0 font-medium text-slate-700">{String(value)}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {initial && (
          <div className="mt-4 flex items-center justify-between rounded-lg bg-slate-50 px-3 py-2.5 text-xs">
            <span className="flex items-center gap-1.5 text-slate-500">
              {initial.passwordHash ? (
                <><CheckCircle2 size={13} className="text-emerald-500" /> Account created — password is set</>
              ) : (
                <><XCircle size={13} className="text-amber-500" /> No account yet — hasn't set a password</>
              )}
            </span>
            {initial.passwordHash && (
              <button
                onClick={() => setShowResetConfirm(true)}
                className="font-medium text-red-500 hover:underline"
              >
                Reset password
              </button>
            )}
          </div>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Cancel
          </button>
          <button
            disabled={!canSave}
            onClick={() => canSave && onSave(form)}
            className="rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
          >
            Save
          </button>
        </div>

        {showResetConfirm && (
          <ConfirmModal
            title="Reset this participant's password?"
            message={`${initial.name} will be signed out everywhere and need to create their account again with a new password.`}
            confirmLabel="Reset Password"
            onCancel={() => setShowResetConfirm(false)}
            onConfirm={() => {
              updateParticipant(initial.employeeId, { passwordHash: null, passwordSalt: null });
              setShowResetConfirm(false);
            }}
          />
        )}
      </div>
    </div>
  );
}

const graded = assessments.filter((a) => a.score !== null);
const passedCount = graded.filter((a) => a.status === "Pass").length;
const failedCount = graded.filter((a) => a.status === "Fail").length;
const avgScore = Math.round(graded.reduce((s, a) => s + a.score, 0) / graded.length);
const certCount = assessments.filter((a) => a.certificate).length;
const awaiting = assessments.filter((a) => a.status === "Grading").length;
const latest = [...graded].sort((a, b) => new Date(b.submitted) - new Date(a.submitted))[0];

function StatusBadge({ status }) {
  const map = {
    Pass: "bg-emerald-50 text-emerald-700 ring-emerald-200",
    Fail: "bg-red-50 text-red-600 ring-red-200",
    Grading: "bg-sky-50 text-sky-600 ring-sky-200",
  };
  const dot = { Pass: "bg-emerald-500", Fail: "bg-red-500", Grading: "bg-sky-500" };
  return (
    <span className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-medium ring-1 ${map[status]}`}>
      <span className={`h-1.5 w-1.5 rounded-full ${dot[status]}`} />
      {status}
    </span>
  );
}

// Original monogram — an abstract "M" built from two peaks, on a rounded
// gradient badge with a small accent mark. Not a reproduction of any real
// company's registered logo; it's a standalone mark for this app's brand.
function LogoMark({ size = 36, light = false }) {
  const uid = useId();
  const gradId = `mercerGrad-${uid}`;
  return (
    <svg width={size} height={size} viewBox="0 0 48 48" fill="none" aria-hidden="true">
      <defs>
        <linearGradient id={gradId} x1="0" y1="0" x2="48" y2="48" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor={BRAND_NAVY} />
          <stop offset="55%" stopColor={BRAND_NAVY_MID} />
          <stop offset="100%" stopColor={BRAND_TEAL} />
        </linearGradient>
      </defs>
      <rect x="1" y="1" width="46" height="46" rx="13" fill={`url(#${gradId})`} stroke={light ? "rgba(255,255,255,0.35)" : "rgba(255,255,255,0.15)"} strokeWidth="1" />
      <path d="M12 34V15.5c0-.7.8-1.1 1.4-.6L24 25l10.6-10.1c.6-.5 1.4-.1 1.4.6V34" stroke="white" strokeWidth="3.4" strokeLinecap="round" strokeLinejoin="round" fill="none" />
      <circle cx="37" cy="12" r="3.2" fill="#5eead4" />
    </svg>
  );
}

// Large, very low-opacity monogram used as a corner watermark on hero panels.
function LogoWatermark({ size = 220, className = "", style }) {
  return (
    <img
      src={MERCER_M_MASK_DATA_URL}
      alt=""
      aria-hidden="true"
      width={size}
      height={size}
      className={`pointer-events-none absolute ${className}`}
      style={{ opacity: 0.14, ...style }}
    />
  );
}

function Logo({ light, size = 36, tagline, animated = false }) {
  return (
    <div className="flex items-center gap-2.5">
      {animated && (
        <style>{`
          @keyframes mercerLogoIn {
            0% { opacity: 0; transform: translateY(-8px) scale(0.9); }
            100% { opacity: 1; transform: translateY(0) scale(1); }
          }
          @keyframes mercerLogoPulse {
            0%, 100% { transform: scale(1); }
            50% { transform: scale(1.08); }
          }
          @keyframes mercerLogoGlow {
            0%, 100% { opacity: 0.45; transform: scale(1); }
            50% { opacity: 0.9; transform: scale(1.25); }
          }
        `}</style>
      )}
      <div className="relative" style={animated ? { animation: "mercerLogoIn 700ms ease-out both" } : undefined}>
        {animated && (
          <div
            className="pointer-events-none absolute rounded-full"
            style={{
              inset: -12,
              background: "radial-gradient(circle, rgba(94,234,212,0.65), rgba(255,255,255,0.25) 45%, rgba(255,255,255,0) 75%)",
              animation: "mercerLogoGlow 2.2s ease-in-out infinite",
              filter: "blur(9px)",
            }}
          />
        )}
        <div style={animated ? { animation: "mercerLogoPulse 2.6s ease-in-out infinite", position: "relative" } : { position: "relative" }}>
          <MercerLogoImg size={size} radius={size / 4} />
        </div>
      </div>
      <div className="leading-tight" style={animated ? { animation: "mercerLogoIn 700ms ease-out 120ms both" } : undefined}>
        <div className={`text-lg font-bold tracking-tight ${light ? "text-white" : ""}`} style={light ? undefined : { color: BRAND_NAVY }}>
          Mercer<span style={{ color: BRAND_TEAL }}>Assess</span>
        </div>
        {tagline && <div className={`text-xs ${light ? "text-white/60" : "text-slate-400"}`}>{tagline}</div>}
      </div>
    </div>
  );
}

// Real notifications derived from live data — unanswered surveys matching
// the participant's audience, and upcoming meetings in the next 7 days.
// Nothing is stored separately; this is computed fresh each time the bell
// data loads.
function NotificationBell({ setView }) {
  const { currentUser } = useAppData();
  const { surveys } = useSharedSurveys();
  const [meetings, setMeetings] = useState([]);
  const [responseStatus, setResponseStatus] = useState({});
  const [open, setOpen] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch(`/api/data?key=${encodeURIComponent(ATTENDANCE_STORAGE_KEY)}`);
        if (res.ok) {
          const data = await res.json();
          const parsed = JSON.parse(data.value || "[]");
          if (Array.isArray(parsed)) setMeetings(parsed);
        }
      } catch (err) {
        // nothing saved yet
      }
    })();
  }, []);

  const pendingSurveys = surveys.filter((s) => s.sourceType === "built" && (s.projectId || "default") === (currentUser?.projectId || "default") && audienceMatches(s.audience, currentUser?.cohort));
  const pendingSurveyIds = pendingSurveys.map((s) => s.id).join(",");

  useEffect(() => {
    if (!currentUser || pendingSurveys.length === 0) return;
    let cancelled = false;
    (async () => {
      const results = {};
      for (const s of pendingSurveys) {
        try {
          const res = await fetch(`/api/survey-response?surveyId=${encodeURIComponent(s.id)}`);
          const data = res.ok ? await res.json() : { responses: [] };
          results[s.id] = (data.responses || []).some((r) => r.respondentEmployeeId === currentUser.employeeId);
        } catch (err) {
          results[s.id] = false;
        }
      }
      if (!cancelled) setResponseStatus(results);
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser?.employeeId, pendingSurveyIds]);

  const surveyNotifications = currentUser
    ? pendingSurveys
        .filter((s) => !responseStatus[s.id])
        .map((s) => ({ type: "survey", id: `survey-${s.id}`, text: `New survey to complete: ${s.name}`, onClick: () => { setView("surveys"); setOpen(false); } }))
    : [];

  const today = new Date().toISOString().slice(0, 10);
  const in7Days = new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10);
  const meetingNotifications = currentUser
    ? meetings
        .filter((m) => m.date >= today && m.date <= in7Days && (m.projectId || "default") === (currentUser?.projectId || "default") && audienceMatches(m.audience, currentUser?.cohort))
        .map((m) => ({ type: "meeting", id: `meeting-${m.id}`, text: `Upcoming: ${m.name} on ${m.date}${m.startTime ? ` at ${m.startTime}` : ""}`, onClick: () => setOpen(false) }))
    : [];

  const notifications = [...surveyNotifications, ...meetingNotifications];

  return (
    <div className="relative">
      <button onClick={() => setOpen((o) => !o)} className="relative text-slate-400 hover:text-slate-600">
        <Bell size={19} />
        {notifications.length > 0 && (
          <span
            className="absolute -right-1.5 -top-1.5 flex h-4 w-4 items-center justify-center rounded-full bg-red-500 font-bold text-white ring-2 ring-white"
            style={{ fontSize: 9 }}
          >
            {notifications.length}
          </span>
        )}
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-20" onClick={() => setOpen(false)} />
          <div className="absolute right-0 top-full z-30 mt-2 rounded-xl border border-slate-200 bg-white p-2 shadow-lg" style={{ width: 320 }}>
            <div className="px-2 py-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">Notifications</div>
            {notifications.length === 0 ? (
              <div className="px-2 py-6 text-center text-sm text-slate-400">You're all caught up.</div>
            ) : (
              <div className="space-y-1" style={{ maxHeight: 320, overflowY: "auto" }}>
                {notifications.map((n) => (
                  <button key={n.id} onClick={n.onClick} className="flex w-full items-start gap-2 rounded-lg px-2 py-2 text-left text-sm hover:bg-slate-50">
                    <span className="mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-blue-50 text-blue-600">
                      {n.type === "survey" ? <PieChartIcon size={12} /> : <Calendar size={12} />}
                    </span>
                    <span className="text-slate-700">{n.text}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function TopNav({ view, setView, onLogout, isAdmin }) {
  const { currentUser } = useAppData();
  const items = [
    { key: "dashboard", label: "Dashboard", icon: Home, requiresUser: true },
    { key: "profile", label: "My Profile", icon: User, requiresUser: true },
    { key: "surveys", label: "Surveys", icon: ClipboardList, requiresUser: true },
    { key: "myGroup", label: "My Group", icon: Users, requiresUser: true },
    { key: "myMeetings", label: "Meetings", icon: Calendar, requiresUser: true },
    { key: "myResources", label: "Resources", icon: FileText, requiresUser: true },
    { key: "survey", label: "Survey Analytics", icon: PieChartIcon, requiresAdmin: true },
    { key: "admin", label: "Admin", icon: ShieldCheck, requiresAdmin: true },
  ].filter((item) => (!item.requiresUser || currentUser) && (!item.requiresAdmin || isAdmin));
  return (
    <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/95 backdrop-blur px-6 py-3">
      <div className="mx-auto flex max-w-7xl items-center justify-between">
        <button onClick={() => setView(currentUser ? "dashboard" : "admin")} className="shrink-0">
          <Logo />
        </button>
        <nav className="flex items-center gap-6">
          {items.map(({ key, label, icon: Icon }) => (
            <button
              key={key}
              onClick={() => setView(key)}
              className={`flex items-center gap-1.5 text-sm font-medium transition-colors ${
                view === key ? "text-blue-700" : "text-slate-500 hover:text-slate-700"
              }`}
            >
              <Icon size={17} strokeWidth={2} />
              {label}
            </button>
          ))}
        </nav>
        <div className="flex items-center gap-4">
          <NotificationBell setView={setView} />
          <div className="flex items-center gap-2.5">
            <div className="flex h-8 w-8 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">
              {currentUser ? initialsFromName(currentUser.name) : <ShieldCheck size={14} />}
            </div>
            <div className="leading-tight">
              <div className="text-sm font-semibold text-slate-800">{currentUser ? currentUser.name : "Admin"}</div>
              <div className="text-xs text-slate-400">{currentUser ? currentUser.cohort : "Administrator"}</div>
            </div>
          </div>
          <button onClick={onLogout} title="Log out" className="text-slate-400 hover:text-red-500">
            <LogOut size={17} />
          </button>
        </div>
      </div>
    </header>
  );
}

function StatCard({ icon: Icon, label, value, sub, iconBg, iconColor }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-5">
      <div className="mb-3 flex items-center justify-between">
        <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">{label}</span>
        <span className={`flex h-8 w-8 items-center justify-center rounded-full ${iconBg}`}>
          <Icon size={15} className={iconColor} />
        </span>
      </div>
      <div className="text-3xl font-bold text-slate-800">{value}</div>
      <div className="mt-1 text-xs text-slate-400">{sub}</div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Dashboard
// ---------------------------------------------------------------------------

function ScoreRing({ value, size = 150, light = false }) {
  const r = 62;
  const c = 2 * Math.PI * r;
  const pct = Math.min(Math.max(value, 0), 100) / 100;
  const color = value >= passingThreshold ? "#0d9488" : "#f97316";
  const trackColor = light ? "rgba(255,255,255,0.25)" : "#e2e8f0";
  const valueColor = light ? "#ffffff" : "#1e293b";
  const labelColor = light ? "rgba(255,255,255,0.7)" : "#94a3b8";
  return (
    <svg width={size} height={size} viewBox="0 0 150 150">
      <circle cx="75" cy="75" r={r} fill="none" stroke={trackColor} strokeWidth="12" />
      <circle
        cx="75" cy="75" r={r} fill="none" stroke={color} strokeWidth="12" strokeLinecap="round"
        strokeDasharray={c} strokeDashoffset={c * (1 - pct)} transform="rotate(-90 75 75)"
      />
      <text x="75" y="70" textAnchor="middle" fontSize="34" fontWeight="700" fill={valueColor}>{value}</text>
      <text x="75" y="90" textAnchor="middle" fontSize="11" fill={labelColor}>/ 100</text>
    </svg>
  );
}

function Dashboard({ openAssessment }) {
  const { currentUser, projects, participants, resendVerificationEmail } = useAppData();
  const workforceAssessment = assessments.find((a) => a.id === "workforce");
  const { reflection: myReflection, loading: reflectionLoading, save: saveReflection } = useMyReflection(workforceAssessment.id, currentUser.employeeId);
  const radarData = latest.criteria.map((c) => ({ criterion: c.name, score: c.score }));
  const myProject = projects.find((p) => p.id === (currentUser.projectId || "default")) || projects[0];
  const myGroup = (myProject?.groups || []).find((g) => g.participantIds.includes(currentUser.employeeId));
  const groupmates = myGroup ? participants.filter((p) => myGroup.participantIds.includes(p.employeeId) && p.employeeId !== currentUser.employeeId) : [];
  const [resendStatus, setResendStatus] = useState(null); // null | "sending" | "sent"

  const handleResendVerification = async () => {
    setResendStatus("sending");
    try {
      await resendVerificationEmail(currentUser.employeeId);
      setResendStatus("sent");
    } catch (err) {
      setResendStatus(null);
    }
  };

  return (
    <div className="mx-auto max-w-7xl space-y-6 px-6 py-6">
      {currentUser.emailVerified === false && (
        <div className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-amber-200 bg-amber-50 px-5 py-4">
          <div className="flex items-center gap-2 text-sm text-amber-800">
            <XCircle size={16} className="shrink-0" />
            Please verify your email — check your inbox for the link from signup. Some actions, like meeting check-in, are unavailable until then.
          </div>
          <button
            onClick={handleResendVerification}
            disabled={resendStatus === "sending" || resendStatus === "sent"}
            className="flex shrink-0 items-center gap-1.5 rounded-lg border border-amber-300 bg-white px-3 py-1.5 text-xs font-semibold text-amber-800 hover:bg-amber-100 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {resendStatus === "sending" ? <RefreshCw size={12} className="animate-spin" /> : <Mail size={12} />}
            {resendStatus === "sent" ? "Email sent!" : "Resend verification email"}
          </button>
        </div>
      )}
      {/* Hero */}
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <div className="absolute -right-10 -top-16 h-56 w-56 rounded-full bg-white/5" />
        <div className="absolute right-40 h-40 w-40 rounded-full bg-white/5" style={{ bottom: -60 }} />
        <LogoWatermark size={200} className="-right-6 -bottom-10" />
        <div className="relative flex flex-wrap items-start justify-between gap-8">
          <div>
            <div className="mb-3 flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-full bg-white/15 text-sm font-semibold">
                {initialsFromName(currentUser.name)}
              </div>
              <span className="text-sm text-white/70">{currentUser.cohort} · {currentUser.dept}</span>
            </div>
            <h1 className="text-3xl font-bold">Welcome back, {currentUser.name.split(" ")[0]}</h1>

            <div className="mt-6 flex flex-wrap gap-3">
              <div className="rounded-xl bg-white/10 px-4 py-3">
                <div className="text-xs text-white/60">Employee ID</div>
                <div className="font-semibold">{currentUser.employeeId}</div>
              </div>
              <div className="rounded-xl bg-white/10 px-4 py-3">
                <div className="text-xs text-white/60">Assessments Passed</div>
                <div className="font-semibold">{passedCount} / {graded.length}</div>
              </div>
              <div className="rounded-xl bg-white/10 px-4 py-3">
                <div className="text-xs text-white/60">Average Score</div>
                <div className="font-semibold">{avgScore}%</div>
              </div>
            </div>

            <div className="mt-4 max-w-sm rounded-xl bg-white/10 px-4 py-3">
              <div className="text-xs text-white/60">Latest Assessment</div>
              <div className="font-semibold">{latest.name}</div>
              <div className="mt-1 flex items-center gap-2 text-sm">
                <span>Score: {latest.score ?? "—"}{latest.score !== null ? "%" : ""}</span>
                <span>·</span>
                <span className={latest.status === "Pass" ? "text-emerald-300" : latest.status === "Fail" ? "text-orange-300" : "text-sky-300"}>
                  {latest.status === "Fail" ? "✗ FAIL" : latest.status === "Pass" ? "✓ PASS" : "● GRADING"}
                </span>
              </div>
            </div>

            <div className="mt-6 flex gap-3">
              <button onClick={() => openAssessment()} className="rounded-lg bg-white px-4 py-2.5 text-sm font-semibold hover:bg-white/90" style={{ color: BRAND_NAVY }}>
                View Full Profile
              </button>
              <button
                onClick={() => document.getElementById("dashboard-certificates")?.scrollIntoView({ behavior: "smooth", block: "start" })}
                className="rounded-lg bg-white/15 px-4 py-2.5 text-sm font-semibold hover:bg-white/25"
              >
                My Certificates
              </button>
            </div>
          </div>

          <div className="flex flex-col items-center rounded-2xl bg-white/10 px-8 py-6">
            <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-white/60">Overall Score</div>
            <ScoreRing value={latest.score ?? 0} light />
            <div className="mt-2 text-sm font-medium">{latest.name}</div>
            <div className="mt-1"><StatusBadge status={latest.status} /></div>
            <div className="mt-1 text-xs text-white/60">Passing threshold: {passingThreshold}%</div>
          </div>
        </div>
      </div>

      {/* Stat cards */}
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <StatCard icon={FileText} label="Total Assessments" value={assessments.length} sub={`${graded.length} graded`} iconBg="bg-blue-50" iconColor="text-blue-600" />
        <StatCard icon={GraduationCap} label="Assessments Passed" value={passedCount} sub={`${failedCount} failed`} iconBg="bg-emerald-50" iconColor="text-emerald-600" />
        <StatCard icon={Sparkles} label="Awaiting Results" value={awaiting} sub="In review" iconBg="bg-amber-50" iconColor="text-amber-600" />
        <StatCard icon={Award} label="Certificates Earned" value={certCount} sub="Available to download" iconBg="bg-teal-50" iconColor="text-teal-600" />
      </div>

      {/* History + radar */}
      <div className="flex flex-wrap items-start gap-6">
        <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "2 1 420px" }}>
          <div className="mb-2 flex items-center justify-between">
            <div>
              <h2 className="text-lg font-bold text-slate-800">Assessment History</h2>
              <p className="text-sm text-slate-400">{assessments.length} assessments enrolled</p>
            </div>
            <button className="flex items-center gap-1 text-sm font-medium text-blue-700 hover:underline">
              View all <ChevronRight size={14} />
            </button>
          </div>
          <div className="divide-y divide-slate-50">
            {assessments.map((a) => (
              <div
                key={a.id}
                onClick={() => openAssessment(a.id)}
                className="-mx-2 flex cursor-pointer flex-wrap items-center justify-between gap-x-4 gap-y-1 rounded-lg px-2 py-3 hover:bg-slate-50"
              >
                <div className="min-w-0">
                  <div className="truncate font-medium text-slate-800">{a.name}</div>
                  <div className="flex flex-wrap items-center gap-x-2 text-xs text-slate-400">
                    <span className="truncate">{a.program}</span>
                    <span>·</span>
                    <span className="whitespace-nowrap">{a.submitted}</span>
                  </div>
                </div>
                <div className="flex shrink-0 items-center gap-3">
                  <span className={`text-sm font-semibold ${a.status === "Pass" ? "text-emerald-600" : a.status === "Fail" ? "text-red-600" : "text-slate-400"}`}>
                    {a.score !== null ? `${a.score}%` : "—"}
                  </span>
                  <StatusBadge status={a.status} />
                  {a.certificate ? (
                    <span title="Certificate available" className="flex items-center gap-1 text-xs font-medium text-teal-600"><Download size={13} /></span>
                  ) : (
                    <span style={{ width: 13, display: "inline-block" }} />
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "1 1 280px" }}>
          <h2 className="text-lg font-bold text-slate-800">Criterion Radar</h2>
          <p className="mb-2 text-sm text-slate-400">{latest.name}</p>
          <ResponsiveContainer width="100%" height={260}>
            <RadarChart data={radarData} outerRadius="75%">
              <PolarGrid stroke="#e2e8f0" />
              <PolarAngleAxis dataKey="criterion" tick={{ fontSize: 11, fill: "#64748b" }} />
              <Radar dataKey="score" stroke="#1d4ed8" fill="#1d4ed8" fillOpacity={0.25} />
            </RadarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Feedback preview + certs */}
      <div className="flex flex-wrap items-start gap-6">
        <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "2 1 420px" }}>
          <div className="mb-4 flex items-center justify-between">
            <div>
              <h2 className="text-lg font-bold text-slate-800">Feedback Preview</h2>
              <p className="text-sm text-slate-400">AI-generated criterion feedback</p>
            </div>
            <button onClick={() => openAssessment(latest.id)} className="flex items-center gap-1 text-sm font-medium text-blue-700 hover:underline">
              Full feedback <ChevronRight size={14} />
            </button>
          </div>
          <div className="mb-4 grid grid-cols-2 gap-4">
            <div className="rounded-xl border border-emerald-100 bg-emerald-50 p-4">
              <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-emerald-600">Strongest Area</div>
              <div className="font-bold text-emerald-800">{latest.feedback?.strongest.name ?? latest.criteria[4].name}</div>
              <div className="text-sm text-emerald-700">{latest.feedback?.strongest.score ?? latest.criteria[4].score} / 100</div>
            </div>
            <div className="rounded-xl border border-amber-100 bg-amber-50 p-4">
              <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-amber-600">Area for Growth</div>
              <div className="font-bold text-amber-800">{latest.feedback?.growth.name ?? latest.criteria[0].name}</div>
              <div className="text-sm text-amber-700">{latest.feedback?.growth.score ?? latest.criteria[0].score} / 100</div>
            </div>
          </div>
          {(latest.feedback?.notes ?? []).map((n) => (
            <div key={n.name} className="mb-3 border-l-2 border-amber-300 pl-3">
              <div className="text-sm font-semibold text-slate-700">{n.name} <span className="text-amber-600">{n.score}</span></div>
              <p className="text-sm text-slate-500">{n.text}</p>
            </div>
          ))}
        </div>

        <div id="dashboard-certificates" className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "1 1 280px" }}>
          <div className="mb-4 flex items-center justify-between">
            <h2 className="text-lg font-bold text-slate-800">My Certificates</h2>
            <span className="text-sm text-slate-400">{certCount} earned</span>
          </div>
          {assessments.filter((a) => a.certificate).map((a) => (
            <div key={a.id} className="mb-3 overflow-hidden rounded-xl border border-slate-200">
              <div className="relative overflow-hidden px-4 py-4 text-center text-white" style={brandGradient()}>
                <LogoWatermark size={90} className="-right-4 -top-4" />
                <div className="relative mx-auto mb-1.5 flex h-9 w-9 items-center justify-center">
                  <MercerLogoImg size={30} radius={8} />
                </div>
                <div className="relative text-sm font-semibold">Certificate of Completion</div>
                <div className="relative text-xs text-white/80">{a.name}</div>
              </div>
              <div className="flex items-center justify-between gap-2 px-4 py-3">
                <div>
                  <div className="text-sm font-medium text-slate-700">{a.name}</div>
                  <div className="text-xs text-slate-400">Score: <span className="font-semibold text-emerald-600">{a.score}%</span></div>
                </div>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => openLinkedInAddToProfile({
                      programName: a.name,
                      score: a.score,
                      certId: `MC-${currentUser.employeeId}-${a.id}`,
                      issueDate: new Date().toISOString().slice(0, 10),
                    })}
                    title="Add this certification to your LinkedIn profile"
                    className="flex items-center gap-1 rounded-lg px-3 py-1.5 text-xs font-semibold text-white hover:opacity-90"
                    style={{ backgroundColor: "#0a66c2" }}
                  >
                    <Share2 size={12} /> LinkedIn
                  </button>
                  <button
                    onClick={() => generateCertificateImage({
                      participantName: currentUser.name,
                      programName: a.name,
                      date: new Date().toISOString().slice(0, 10),
                      certId: `MC-${currentUser.employeeId}-${a.id}`,
                    })}
                    className="flex items-center gap-1 rounded-lg bg-teal-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-teal-700"
                  >
                    <Download size={12} /> Image
                  </button>
                </div>
              </div>
            </div>
          ))}
        </div>

        {myGroup && (
          <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "1 1 280px" }}>
            <div className="mb-4 flex items-center justify-between">
              <h2 className="text-lg font-bold text-slate-800">My Group</h2>
              <span className="rounded-full bg-blue-50 px-2.5 py-1 text-xs font-semibold text-blue-700">{myGroup.name}</span>
            </div>
            {groupmates.length === 0 ? (
              <p className="text-sm text-slate-400">No other participants in your group yet.</p>
            ) : (
              <div className="space-y-2">
                {groupmates.map((p) => (
                  <div key={p.employeeId} className="flex items-center gap-2.5 rounded-lg bg-slate-50 px-3 py-2">
                    <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">{initialsFromName(p.name)}</div>
                    <div className="min-w-0">
                      <div className="text-sm font-medium text-slate-700">{p.name}</div>
                      <div className="truncate text-xs text-slate-400">{p.cohort} · <CopyEmailButton email={p.email} className="text-blue-600 hover:underline" /></div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>

      {/* AI Reflection Grading */}
      {reflectionLoading ? (
        <div className="flex items-center justify-center rounded-2xl border border-slate-200 bg-white py-10 text-slate-400">
          <RefreshCw size={20} className="animate-spin" />
        </div>
      ) : myReflection ? (
        <div className="space-y-4">
          <AIGradingPanel assessment={{ ...workforceAssessment, reflection: myReflection.text, submitted: myReflection.submittedAt.slice(0, 10) }} />
          <details className="rounded-xl border border-slate-100 bg-slate-50 px-4 py-3 text-sm text-slate-500">
            <summary className="cursor-pointer font-medium text-slate-600">Update your submitted reflection</summary>
            <div className="mt-3">
              <ReflectionSubmissionPanel assessmentName={workforceAssessment.name} existing={myReflection} onSaved={saveReflection} />
            </div>
          </details>
        </div>
      ) : (
        <ReflectionSubmissionPanel assessmentName={workforceAssessment.name} existing={null} onSaved={saveReflection} />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Reflection submission — participants paste their own text or upload a
// Word/PDF file (extracted to text server-side), replacing the mock
// reflection text that used to be hardcoded here.
// ---------------------------------------------------------------------------

function useMyReflection(assessmentId, employeeId) {
  const [reflection, setReflection] = useState(null); // { text, fileName, submittedAt } | null
  const [loading, setLoading] = useState(true);
  const storageKey = `reflection-${assessmentId}-${employeeId}`;

  const load = async () => {
    setLoading(true);
    try {
      const result = await storage.get(storageKey);
      if (result && result.value) {
        setReflection(JSON.parse(result.value));
      } else {
        setReflection(null);
      }
    } catch (err) {
      setReflection(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [assessmentId, employeeId]);

  const save = async (text, fileName) => {
    const entry = { text, fileName: fileName || null, submittedAt: new Date().toISOString() };
    await storage.set(storageKey, JSON.stringify(entry));
    setReflection(entry);
  };

  return { reflection, loading, save, refresh: load };
}

function ReflectionSubmissionPanel({ assessmentName, existing, onSaved }) {
  const [mode, setMode] = useState("paste"); // 'paste' | 'upload'
  const [text, setText] = useState(existing?.text || "");
  const [fileName, setFileName] = useState(existing?.fileName || "");
  const [extracting, setExtracting] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = React.useRef(null);

  const handleFile = async (file) => {
    if (!file) return;
    setError(null);
    setExtracting(true);
    setFileName(file.name);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const base64 = btoa(new Uint8Array(arrayBuffer).reduce((s, b) => s + String.fromCharCode(b), ""));
      const res = await fetch("/api/extract-text", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fileName: file.name, mimeType: file.type, fileData: base64 }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to extract text from file.");
      setText(data.text);
    } catch (err) {
      setError(err.message || "Failed to extract text from file.");
    } finally {
      setExtracting(false);
    }
  };

  const handleSubmit = async () => {
    if (!text.trim()) return;
    setSaving(true);
    setError(null);
    try {
      await onSaved(text.trim(), mode === "upload" ? fileName : null);
    } catch (err) {
      setError("Failed to save your reflection. Please try again.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-6">
      <div className="mb-1 flex items-center gap-2">
        <FileText size={18} className="text-blue-600" />
        <h2 className="text-lg font-bold text-slate-800">Submit Your Reflection</h2>
      </div>
      <p className="mb-4 text-sm text-slate-400">
        {existing ? `Last submitted ${new Date(existing.submittedAt).toLocaleDateString()} — you can update it below.` : `Write a short reflection for "${assessmentName}" to get AI feedback.`}
      </p>

      <div className="mb-4 flex rounded-lg border border-slate-200 p-1" style={{ maxWidth: 280 }}>
        <button
          type="button"
          onClick={() => setMode("paste")}
          className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-1.5 text-xs font-medium ${mode === "paste" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-slate-50"}`}
        >
          <Pencil size={12} /> Paste text
        </button>
        <button
          type="button"
          onClick={() => setMode("upload")}
          className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-1.5 text-xs font-medium ${mode === "upload" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-slate-50"}`}
        >
          <UploadCloud size={12} /> Upload file
        </button>
      </div>

      {mode === "upload" && (
        <>
          <input
            ref={fileInputRef}
            type="file"
            accept=".docx,.pdf"
            className="hidden"
            onChange={(e) => handleFile(e.target.files && e.target.files[0])}
          />
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              const file = e.dataTransfer.files && e.dataTransfer.files[0];
              if (file) handleFile(file);
            }}
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
            className={`mb-3 flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-4 py-6 text-center transition-colors ${
              dragOver ? "border-blue-400 bg-blue-50" : "border-slate-200 bg-slate-50"
            }`}
          >
            {extracting ? (
              <>
                <RefreshCw size={20} className="mb-2 animate-spin text-blue-500" />
                <p className="text-sm font-medium text-slate-600">Reading {fileName}…</p>
              </>
            ) : (
              <>
                <UploadCloud size={20} className={dragOver ? "text-blue-500" : "text-slate-400"} />
                <p className="mt-2 text-sm font-medium text-slate-600">{fileName || "Drag & drop a .docx or .pdf file, or click to browse"}</p>
              </>
            )}
          </div>
        </>
      )}

      <textarea
        rows={6}
        value={text}
        onChange={(e) => setText(e.target.value)}
        placeholder="Your reflection…"
        className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
      />
      {mode === "upload" && text && (
        <p className="mt-1 text-xs text-slate-400">Extracted text shown above — feel free to edit it before submitting.</p>
      )}

      {error && <div className="mt-3 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{error}</div>}

      <button
        disabled={!text.trim() || saving || extracting}
        onClick={handleSubmit}
        className="mt-4 flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
      >
        {saving ? <RefreshCw size={14} className="animate-spin" /> : <CheckCircle2 size={14} />}
        {existing ? "Update Reflection" : "Submit Reflection"}
      </button>
    </div>
  );
}

function AIGradingPanel({ assessment }) {
  const { currentUser, projects, currentProjectId, updateParticipant } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId) || projects[0];
  const reflectionCriterion = currentProject?.criteria?.find((c) => c.isReflectionCriterion);

  const [openRow, setOpenRow] = useState(null);
  const [aiCriteria, setAiCriteria] = useState(null); // null until a real grade comes back
  const [grading, setGrading] = useState(false);
  const [gradeError, setGradeError] = useState(null);

  // This internal rubric (Innovation, Clarity, etc.) is how the AI evaluates
  // the reflection's writing quality — a richer breakdown than a single
  // number. Its own weighted result becomes the score for whichever ONE
  // criterion the project designated as "AI-graded from reflections" (see
  // Grading Criteria), not the participant's whole program grade.
  const displayCriteria = assessment.criteria.map((c) => {
    const aiMatch = aiCriteria && aiCriteria.find((a) => a.name === c.name);
    return aiMatch ? { ...c, score: aiMatch.score, feedback: aiMatch.feedback } : c;
  });
  const reflectionTotal = displayCriteria.reduce((s, c) => s + (c.score * c.weight) / 100, 0);
  const isGraded = Boolean(aiCriteria);

  // The real overall grade is the weighted total across every criterion the
  // project actually tracks (attendance, ALPs, modules, this reflection,
  // etc.) — not just the reflection's own internal breakdown.
  const projectTotal = computeWeightedTotal(currentUser, currentProject);
  const displayedTotal = projectTotal ?? reflectionTotal;
  const passed = displayedTotal >= passingThreshold;

  const handleGrade = async () => {
    setGrading(true);
    setGradeError(null);
    try {
      const res = await fetch("/api/grade-reflection", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          reflection: assessment.reflection,
          criteria: assessment.criteria.map((c) => ({ name: c.name, weight: c.weight })),
          assessmentName: assessment.name,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        setGradeError(data.error || "Failed to grade reflection.");
      } else {
        setAiCriteria(data.criteria);
        if (reflectionCriterion) {
          const newTotal = data.criteria.reduce((s, c) => s + (c.score * c.weight) / 100, 0);
          updateParticipant(currentUser.employeeId, {
            criteriaScores: { ...(currentUser.criteriaScores || {}), [reflectionCriterion.name]: Math.round(newTotal * 10) / 10 },
          });
        }
      }
    } catch (err) {
      setGradeError("Network error while grading. Please try again.");
    } finally {
      setGrading(false);
    }
  };

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-6">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <Sparkles size={18} className="text-blue-600" />
          <div>
            <h2 className="text-lg font-bold text-slate-800">AI Reflection Grading</h2>
            <p className="text-sm text-slate-400">Automated scoring against weighted criteria — powered by Gemini</p>
          </div>
        </div>
        {isGraded ? (
          <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-medium text-emerald-700 ring-1 ring-emerald-200">
            ✓ Graded by Gemini
          </span>
        ) : (
          <span className="rounded-full bg-amber-50 px-3 py-1 text-xs font-medium text-amber-600 ring-1 ring-amber-200">
            ● Not yet graded — showing placeholder scores
          </span>
        )}
      </div>

      {reflectionCriterion ? (
        <div className="mb-4 flex items-center gap-1.5 rounded-lg bg-blue-50 px-3 py-2 text-xs text-blue-700">
          <Sparkles size={13} /> This grades the <strong>"{reflectionCriterion.name}"</strong> criterion ({reflectionCriterion.weight}% of the total grade) for {currentProject?.name}.
        </div>
      ) : (
        <div className="mb-4 flex items-center gap-1.5 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-700">
          <XCircle size={13} /> No criterion in {currentProject?.name} is marked "AI-graded from reflections" yet — grading will show feedback but won't be saved to a criterion until you set one in Grading Criteria.
        </div>
      )}

      <div className="flex flex-wrap items-start gap-4">
        <div className="rounded-xl border border-slate-100 bg-slate-50 p-4" style={{ flex: "2 1 380px" }}>
          <div className="mb-2 flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-slate-400">
            <FileText size={13} /> Submitted Reflection — {assessment.name}
          </div>
          <p className="text-sm italic text-slate-600">"{assessment.reflection}"</p>
          <div className="mt-2 text-xs text-slate-400">Submitted {assessment.submitted} · Graded {assessment.graded}</div>
          <button
            onClick={handleGrade}
            disabled={grading}
            className="mt-3 flex items-center gap-1.5 rounded-lg bg-blue-800 px-3 py-1.5 text-xs font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {grading ? <RefreshCw size={12} className="animate-spin" /> : <Sparkles size={12} />}
            {grading ? "Grading…" : isGraded ? "Re-grade with AI" : "Grade with AI"}
          </button>
          {gradeError && (
            <div className="mt-2 rounded-lg bg-red-50 px-3 py-2 text-xs text-red-600">
              {gradeError}
              {String(gradeError).includes("GEMINI_API_KEY") && (
                <div className="mt-1">Make sure <code>GEMINI_API_KEY</code> is set in this project's Vercel environment variables, then redeploy.</div>
              )}
            </div>
          )}
        </div>
        <div className="flex flex-col items-center justify-center rounded-xl border border-emerald-100 bg-emerald-50 p-4 text-center" style={{ flex: "1 1 220px" }}>
          <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Overall Weighted Score</div>
          <div className="text-4xl font-bold text-emerald-700">{displayedTotal.toFixed(1)}</div>
          <div className="text-xs text-slate-400">out of 100 {projectTotal !== null ? `· ${currentProject?.name}` : "· reflection only"}</div>
          <span className="mt-2 rounded-full bg-emerald-100 px-3 py-1 text-xs font-semibold text-emerald-700">
            {passed ? "✓ PASSED" : "✗ FAILED"}
          </span>
          <div className="mt-1 text-xs text-slate-400">Passing threshold: {passingThreshold}</div>
        </div>
      </div>

      <div className="mt-5">
        <div className="mb-2 flex items-center gap-2 text-sm font-semibold text-slate-700">
          <BarChart3 size={15} /> Per-Criterion Breakdown <span className="font-normal text-slate-400">(click a row to expand feedback)</span>
        </div>
        <div className="overflow-hidden rounded-xl border border-slate-100">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-400">
                <th className="px-4 py-2 font-semibold">Criterion</th>
                <th className="px-4 py-2 font-semibold">Weight</th>
                <th className="px-4 py-2 font-semibold">Score</th>
                <th className="px-4 py-2 font-semibold">Score Bar</th>
                <th className="px-4 py-2 font-semibold">Weighted</th>
              </tr>
            </thead>
            <tbody>
              {displayCriteria.map((c) => {
                const weighted = (c.score * c.weight) / 100;
                const isOpen = openRow === c.name;
                return (
                  <React.Fragment key={c.name}>
                    <tr onClick={() => setOpenRow(isOpen ? null : c.name)} className="cursor-pointer border-t border-slate-100 hover:bg-slate-50">
                      <td className="flex items-center gap-1 px-4 py-3 font-medium text-slate-700">
                        {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />} {c.name}
                      </td>
                      <td className="px-4 py-3"><span className="rounded bg-slate-100 px-2 py-0.5 text-xs">{c.weight}%</span></td>
                      <td className="px-4 py-3 font-semibold text-teal-700">{c.score} / 100</td>
                      <td className="px-4 py-3">
                        <div className="h-2 w-40 rounded-full bg-slate-100">
                          <div className="h-2 rounded-full bg-teal-500" style={{ width: `${c.score}%` }} />
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-semibold text-blue-700">+{weighted.toFixed(1)} pts</span>
                      </td>
                    </tr>
                    {isOpen && (
                      <tr className="border-t border-slate-50 bg-slate-50/60">
                        <td colSpan={5} className="px-4 py-3 text-sm text-slate-500">
                          {c.feedback ? (
                            <>
                              <span className="font-medium text-slate-700">{c.name}:</span> {c.feedback}
                            </>
                          ) : (
                            <>Placeholder feedback for <span className="font-medium text-slate-700">{c.name}</span>: reflects a score of {c.score}/100 at a {c.weight}% weighting toward the final grade. Click "Grade with AI" above for real feedback.</>
                          )}
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="mt-3 flex items-center justify-between rounded-lg bg-slate-50 px-4 py-2.5 text-sm">
          <span className="text-slate-500">
            {displayCriteria.map((c) => `(${c.score} × ${c.weight}%)`).join(" + ")}
          </span>
          <span className="font-semibold text-slate-700">Reflection sub-score: <span className="text-blue-700">{reflectionTotal.toFixed(1)} / 100</span></span>
        </div>
        {reflectionCriterion && (
          <p className="mt-2 text-xs text-slate-400">This feeds into "{reflectionCriterion.name}" ({reflectionCriterion.weight}% of the overall grade shown above), alongside {currentProject?.name}'s other criteria.</p>
        )}
        {!isGraded && (
          <div className="mt-3 rounded-lg bg-blue-50 px-4 py-2.5 text-xs text-blue-700">
            These are <strong>placeholder scores</strong>, not real grades. Click "Grade with AI" above to have Gemini actually evaluate this reflection against your criteria and weights.
          </div>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Profile
// ---------------------------------------------------------------------------

function Profile() {
  const { currentUser, updateParticipant } = useAppData();
  const [selected, setSelected] = useState(assessments[1].id);
  const [editing, setEditing] = useState(false);
  const current = assessments.find((a) => a.id === selected);
  const trend = graded.map((a) => ({ name: a.name, score: a.score }));

  return (
    <div className="mx-auto max-w-7xl space-y-6 px-6 py-6">
      <div className="overflow-hidden rounded-3xl border border-slate-200">
        <div className="relative h-32 overflow-hidden" style={brandGradient("to right")}>
          <LogoWatermark size={170} className="-right-6 -top-10" />
        </div>
        <div className="relative bg-white px-6 pb-6" style={{ paddingTop: 44 }}>
          <div
            className="absolute flex items-center justify-center rounded-2xl border-4 border-white bg-blue-800 text-xl font-bold text-white"
            style={{ top: -40, left: 24, width: 80, height: 80 }}
          >
            {initialsFromName(currentUser.name)}
          </div>

          <div className="flex flex-wrap items-start justify-between gap-4" style={{ paddingLeft: 96 }}>
            <div>
              <h1 className="text-2xl font-bold text-slate-800">{currentUser.name}</h1>
              <div className="flex flex-wrap items-center gap-2 text-sm text-slate-500">
                <span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-medium text-blue-700">{currentUser.cohort}</span>
                <span>· {currentUser.dept} · {currentUser.employeeId}</span>
              </div>
            </div>
            <div className="flex gap-2">
              <button onClick={() => setEditing(true)} className="flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"><Pencil size={13} /> Edit Profile</button>
              <button className="flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"><Share2 size={13} /> Share</button>
            </div>
          </div>

          <div className="mt-4 flex flex-wrap items-center justify-between gap-4">
            <div className="flex flex-wrap gap-4 text-sm text-slate-500">
              <span>✉ {currentUser.email}</span>
              <span>🗓 Enrolled {currentUser.enrolled}</span>
            </div>
            <div className="flex flex-wrap gap-6">
              <div className="text-center">
                <div className="text-lg font-bold text-slate-800">{assessments.length}</div>
                <div className="text-xs text-slate-400">Assessments</div>
              </div>
              <div className="text-center">
                <div className="text-lg font-bold text-emerald-600">{passedCount}</div>
                <div className="text-xs text-slate-400">Passed</div>
              </div>
              <div className="text-center">
                <div className="text-lg font-bold text-blue-600">{avgScore}%</div>
                <div className="text-xs text-slate-400">Avg Score</div>
              </div>
              <div className="text-center">
                <div className="text-lg font-bold text-teal-600">{certCount}</div>
                <div className="text-xs text-slate-400">Certificates</div>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="flex flex-wrap items-start gap-6">
        <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "2 1 420px" }}>
          <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-lg font-bold text-slate-800">Score Overview</h2>
              <p className="text-sm text-slate-400">Select an assessment to view its score details</p>
            </div>
            <div className="flex gap-2">
              {assessments.map((a) => (
                <button
                  key={a.id}
                  onClick={() => setSelected(a.id)}
                  className={`rounded-lg px-3 py-1.5 text-xs font-medium ${selected === a.id ? "bg-blue-800 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}
                >
                  {a.name.split(" ").slice(0, 2).join(" ")}
                </button>
              ))}
            </div>
          </div>

          {current.score !== null ? (
            <>
              <div className="flex flex-wrap items-center gap-6">
                <ScoreRing value={current.score} size={130} />
                <div>
                  <div className="text-xl font-bold text-slate-800">{current.name}</div>
                  <div className="text-sm text-slate-400">{current.program}</div>
                  <div className="mt-2 flex gap-6 text-sm">
                    <div><div className="text-xs text-slate-400">Submitted</div><div className="font-medium">{current.submitted}</div></div>
                    <div><div className="text-xs text-slate-400">Graded</div><div className="font-medium">{current.graded}</div></div>
                    <div className="rounded-lg bg-emerald-50 px-3 py-1"><div className="text-xs text-emerald-600">Final Score</div><div className="font-bold text-emerald-700">{current.score}%</div></div>
                  </div>
                </div>
              </div>
              <div className={`mt-4 rounded-lg px-4 py-2.5 text-sm ${current.status === "Pass" ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>
                {current.status === "Pass"
                  ? `${current.score - passingThreshold} points above passing threshold — well done!`
                  : `${passingThreshold - current.score} points below passing threshold.`} {current.certificate && "Certificate has been issued and is available for download."}
              </div>

              <table className="mt-5 w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
                    <th className="py-2 font-semibold">Criterion</th>
                    <th className="py-2 font-semibold">Weight</th>
                    <th className="py-2 font-semibold">Raw Score</th>
                    <th className="py-2 font-semibold">Weighted</th>
                    <th className="py-2 font-semibold">Status</th>
                  </tr>
                </thead>
                <tbody>
                  {current.criteria.map((c) => (
                    <tr key={c.name} className="border-b border-slate-50 last:border-0">
                      <td className="py-2.5 font-medium text-slate-700">{c.name}</td>
                      <td className="py-2.5"><span className="rounded bg-slate-100 px-2 py-0.5 text-xs">{c.weight}%</span></td>
                      <td className="py-2.5 font-semibold text-emerald-600">{c.score}</td>
                      <td className="py-2.5">{((c.score * c.weight) / 100).toFixed(1)}</td>
                      <td className="py-2.5"><StatusBadge status={c.score >= passingThreshold ? "Pass" : "Fail"} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          ) : (
            <div className="rounded-lg bg-sky-50 px-4 py-6 text-center text-sm text-sky-600">
              This assessment is still being graded. Check back soon.
            </div>
          )}
        </div>

        <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "1 1 280px" }}>
          <div className="mb-4 flex items-center justify-between">
            <h2 className="text-lg font-bold text-slate-800">Submission Timeline</h2>
          </div>
          <div className="space-y-4">
            {[...assessments].sort((a, b) => new Date(b.submitted) - new Date(a.submitted)).map((a) => (
              <div key={a.id} className="flex gap-3 border-b border-slate-50 pb-4 last:border-0">
                <div className={`mt-1 h-2.5 w-2.5 shrink-0 rounded-full ${a.status === "Pass" ? "bg-emerald-500" : a.status === "Fail" ? "bg-red-500" : "bg-amber-400"}`} />
                <div className="flex-1">
                  <div className="flex items-center justify-between">
                    <span className="text-sm font-semibold text-slate-700">{a.name}</span>
                    {a.score !== null && <span className={`text-sm font-semibold ${a.status === "Pass" ? "text-emerald-600" : "text-red-600"}`}>{a.score}%</span>}
                  </div>
                  <div className="text-xs text-slate-400">{a.program}</div>
                  <div className="mt-1 flex items-center gap-2 text-xs text-slate-400">
                    Submitted {a.submitted}
                    <StatusBadge status={a.status} />
                    {a.certificate && <span className="text-teal-600">🎓 Certified</span>}
                  </div>
                </div>
              </div>
            ))}
          </div>
          <div className="mt-4 flex items-center gap-2 rounded-lg bg-slate-50 px-3 py-2 text-xs text-slate-500">
            <Building2 size={14} /> Enrolled since {currentUser.enrolled} · {currentUser.cohort} · {currentUser.dept}
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        <h2 className="mb-4 text-lg font-bold text-slate-800">Score Trend Across Graded Assessments</h2>
        <ResponsiveContainer width="100%" height={240}>
          <LineChart data={trend}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
            <XAxis dataKey="name" tick={{ fontSize: 11, fill: "#64748b" }} tickFormatter={(v) => v.split(" ")[0]} />
            <YAxis domain={[50, 100]} tick={{ fontSize: 11, fill: "#64748b" }} />
            <Tooltip />
            <ReferenceLine y={passingThreshold} stroke="#f97316" strokeDasharray="4 4" label={{ value: "Pass", fontSize: 10, fill: "#f97316" }} />
            <Line type="monotone" dataKey="score" stroke="#1d4ed8" strokeWidth={3} dot={{ r: 5, fill: "#1d4ed8" }} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      {editing && (
        <ParticipantFormModal
          title="Edit Profile"
          initial={currentUser}
          showStatus={false}
          onClose={() => setEditing(false)}
          onSave={(data) => {
            updateParticipant(currentUser.employeeId, data);
            setEditing(false);
          }}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Participant Surveys section — lets a logged-in participant see and answer
// built surveys from within the app itself, instead of only via QR/link.
// Unlike the anonymous QR flow, responses submitted here are tagged with the
// participant's identity so completion can be tracked.
// ---------------------------------------------------------------------------

function SurveyResponseForm({ survey, onDone, onCancel }) {
  const { currentUser, showToast } = useAppData();
  const [answers, setAnswers] = useState({});
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState(null);

  const setAnswer = (qid, value) => setAnswers((a) => ({ ...a, [qid]: value }));
  const canSubmit = survey.questions.every((q) => (q.type === "rating" ? typeof answers[q.id] === "number" : true));

  const handleSubmit = async () => {
    setSubmitting(true);
    setError(null);
    try {
      const res = await fetch(`/api/survey-response?surveyId=${encodeURIComponent(survey.id)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ answers, respondentEmployeeId: currentUser.employeeId, respondentName: currentUser.name }),
      });
      if (!res.ok) throw new Error("failed");
      showToast(`Thanks! Your response to "${survey.name}" was submitted.`);
      onDone();
    } catch (err) {
      setError("Something went wrong submitting your response. Please try again.");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white">
      <div className="relative flex items-center justify-between overflow-hidden p-5 text-white" style={brandGradient()}>
        <LogoWatermark size={140} className="-right-6" style={{ top: -30 }} />
        <div className="relative flex items-center gap-2.5">
          {survey.logoUrl ? <ClientLogo logoUrl={survey.logoUrl} size={28} radius={6} /> : <PieChartIcon size={20} />}
          <h2 className="text-lg font-bold">{survey.name}</h2>
        </div>
        <button onClick={onCancel} className="relative text-white/70 hover:text-white"><X size={18} /></button>
      </div>
      <div className="p-6">
      <div className="space-y-5">
        {survey.questions.map((q, i) => (
          <div key={q.id}>
            <label className="mb-2 block text-sm font-medium text-slate-700">{i + 1}. {q.text}</label>
            {q.type === "rating" ? (
              <RatingScaleInput value={answers[q.id]} onChange={(n) => setAnswer(q.id, n)} />
            ) : (
              <textarea
                rows={3}
                value={answers[q.id] || ""}
                onChange={(e) => setAnswer(q.id, e.target.value)}
                placeholder="Your answer…"
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
              />
            )}
          </div>
        ))}
      </div>
      {error && <div className="mt-4 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{error}</div>}
      <div className="mt-6 flex justify-end gap-2">
        <button onClick={onCancel} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
          Cancel
        </button>
        <button
          disabled={!canSubmit || submitting}
          onClick={handleSubmit}
          className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
        >
          {submitting ? <RefreshCw size={14} className="animate-spin" /> : <CheckCircle2 size={14} />}
          Submit
        </button>
      </div>
      </div>
    </div>
  );
}

function SurveyListCard({ survey, currentUser, onStart }) {
  const { hasResponded, loading } = useHasResponded(survey.id, currentUser.employeeId);
  return (
    <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-slate-100 bg-white p-4">
      <div>
        <div className="font-medium text-slate-800">{survey.name}</div>
        <div className="text-xs text-slate-400">{survey.questions.length} question{survey.questions.length === 1 ? "" : "s"} · {formatAudience(survey.audience)}</div>
      </div>
      {loading ? (
        <RefreshCw size={16} className="animate-spin text-slate-300" />
      ) : hasResponded ? (
        <span className="flex items-center gap-1.5 rounded-full bg-emerald-50 px-3 py-1.5 text-xs font-medium text-emerald-700">
          <CheckCircle2 size={13} /> Completed
        </span>
      ) : (
        <button onClick={() => onStart(survey)} className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-3 py-1.5 text-sm font-semibold text-white hover:bg-blue-900">
          <PieChartIcon size={13} /> Start Survey
        </button>
      )}
    </div>
  );
}

function ParticipantSurveysPage() {
  const { currentUser } = useAppData();
  const { surveys, loading } = useSharedSurveys();
  const [activeSurvey, setActiveSurvey] = useState(null);
  const [refreshKey, setRefreshKey] = useState(0);

  const available = surveys.filter((s) => s.sourceType === "built" && (s.projectId || "default") === (currentUser.projectId || "default") && audienceMatches(s.audience, currentUser.cohort));

  if (activeSurvey) {
    return (
      <div className="fixed inset-0 z-40 overflow-y-auto p-6" style={brandGradient()}>
        <LogoWatermark size={280} className="-right-16" style={{ bottom: -80 }} />
        <div className="relative mx-auto max-w-2xl space-y-6 py-6">
          <SurveyResponseForm
            survey={activeSurvey}
            onCancel={() => setActiveSurvey(null)}
            onDone={() => { setActiveSurvey(null); setRefreshKey((k) => k + 1); }}
          />
        </div>
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-3xl space-y-6 px-6 py-6">
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <LogoWatermark size={200} className="-right-8 -bottom-12" />
        <div className="flex items-center gap-4">
          <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/15"><PieChartIcon size={22} /></div>
          <div>
            <h1 className="text-2xl font-bold">Surveys</h1>
            <p className="text-sm text-white/70">Surveys available to you — your responses help improve future programs</p>
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        {loading ? (
          <div className="flex items-center justify-center py-10 text-slate-400"><RefreshCw size={20} className="animate-spin" /></div>
        ) : available.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-14 text-center text-slate-400">
            <PieChartIcon size={26} className="mb-2" />
            <p className="text-sm">No surveys available for you right now.</p>
          </div>
        ) : (
          <div key={refreshKey} className="space-y-3">
            {available.map((s) => (
              <SurveyListCard key={s.id} survey={s} currentUser={currentUser} onStart={setActiveSurvey} />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Participant Meetings — a simplified, view-only version of the admin's
// Meetings tab: just upcoming meetings that apply to this participant, each
// with a button to pull up its QR code full-screen (same view used when an
// admin projects it to a room).
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Participant Resources — documents/videos to review, with a simple mark-
// complete toggle tracked per participant.
// ---------------------------------------------------------------------------

// Determines how to render a resource inline: native PDF/video rendering
// where the browser can do it directly, Microsoft's public Office viewer for
// Word/PowerPoint/Excel files (works for any publicly-reachable URL, no
// account needed), platform embeds for YouTube/Vimeo links, and a plain
// "open in new tab" fallback for anything else (e.g. Google Drive/SharePoint
// links, which vary in whether they allow embedding).
function getResourceViewerKind(resource) {
  const name = (resource.fileName || resource.url || "").toLowerCase();
  if (/\.pdf(\?|$)/.test(name)) return "pdf";
  if (/\.(docx?|pptx?|xlsx?)(\?|$)/.test(name)) return "office";
  if (/\.(mp4|mov|webm|m4v)(\?|$)/.test(name)) return "video";
  if (/youtube\.com|youtu\.be/.test(resource.url || "")) return "youtube";
  if (/vimeo\.com/.test(resource.url || "")) return "vimeo";
  return "external";
}

function getYouTubeEmbedUrl(url) {
  const match = url.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/)([\w-]+)/);
  return match ? `https://www.youtube.com/embed/${match[1]}` : null;
}

function getVimeoEmbedUrl(url) {
  const match = url.match(/vimeo\.com\/(\d+)/);
  return match ? `https://player.vimeo.com/video/${match[1]}` : null;
}

// Vercel Blob (and many other file hosts) deliberately sets
// X-Frame-Options: DENY on every file for security, which silently blocks
// browsers from rendering it directly inside an iframe — the frame just
// stays blank, with no visible error. Fetching the file's bytes ourselves
// and creating a local blob: URL sidesteps this entirely, since the iframe
// then loads already-downloaded, same-origin data rather than making a new
// cross-origin frame request that the file's own headers would block.
function PdfViewer({ url, title }) {
  const [objectUrl, setObjectUrl] = useState(null);
  const [status, setStatus] = useState("loading"); // loading | ready | error

  useEffect(() => {
    let cancelled = false;
    let currentUrl = null;
    (async () => {
      try {
        const res = await fetch(url);
        if (!res.ok) throw new Error("fetch failed");
        const blob = await res.blob();
        if (cancelled) return;
        currentUrl = URL.createObjectURL(blob);
        setObjectUrl(currentUrl);
        setStatus("ready");
      } catch (err) {
        if (!cancelled) setStatus("error");
      }
    })();
    return () => {
      cancelled = true;
      if (currentUrl) URL.revokeObjectURL(currentUrl);
    };
  }, [url]);

  if (status === "loading") {
    return (
      <div className="flex h-full items-center justify-center text-slate-400">
        <RefreshCw size={22} className="animate-spin" />
      </div>
    );
  }

  if (status === "error") {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-3 p-8 text-center text-slate-400">
        <FileText size={28} />
        <p className="text-sm">Couldn't load this PDF for inline preview — it'll open in a new tab instead.</p>
        <a href={url} target="_blank" rel="noreferrer" className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
          <Link2 size={14} /> Open "{title}"
        </a>
      </div>
    );
  }

  return <iframe src={objectUrl} title={title} className="h-full w-full border-0" />;
}

// Microsoft's Office viewer previously had no failure handling at all — if
// the underlying file was unreachable (a broken link, or the simulated
// upload URL used in this chat's preview, which isn't a real file), the
// iframe just stayed blank forever with no indication anything was wrong.
// A quick HEAD request first confirms the file is actually reachable before
// attempting to render it, so an unreachable file surfaces a clear message
// instead of silence.
function OfficeViewer({ url, title }) {
  const [status, setStatus] = useState("loading"); // loading | ready | error

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(url, { method: "HEAD" });
        if (!res.ok) throw new Error("unreachable");
        if (!cancelled) setStatus("ready");
      } catch (err) {
        if (!cancelled) setStatus("error");
      }
    })();
    return () => { cancelled = true; };
  }, [url]);

  if (status === "loading") {
    return (
      <div className="flex h-full items-center justify-center text-slate-400">
        <RefreshCw size={22} className="animate-spin" />
      </div>
    );
  }

  if (status === "error") {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-3 p-8 text-center text-slate-400">
        <FileText size={28} />
        <p className="text-sm">Couldn't reach this file for inline preview — it'll open in a new tab instead.</p>
        <a href={url} target="_blank" rel="noreferrer" className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
          <Link2 size={14} /> Open "{title}"
        </a>
      </div>
    );
  }

  return (
    <iframe
      src={`https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(url)}`}
      title={title}
      className="h-full w-full border-0"
    />
  );
}

function ResourceViewerModal({ resource, onClose }) {
  const kind = getResourceViewerKind(resource);
  const youtubeUrl = kind === "youtube" ? getYouTubeEmbedUrl(resource.url) : null;
  const vimeoUrl = kind === "vimeo" ? getVimeoEmbedUrl(resource.url) : null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 p-4">
      <div className="flex w-full flex-col rounded-2xl bg-white shadow-xl" style={{ maxWidth: 960, height: "85vh" }}>
        <div className="flex items-center justify-between border-b border-slate-100 px-5 py-3">
          <div className="flex items-center gap-2">
            <FileText size={16} className="text-blue-700" />
            <h3 className="text-sm font-bold text-slate-800">{resource.title}</h3>
          </div>
          <div className="flex items-center gap-2">
            <a href={resource.url} target="_blank" rel="noreferrer" className="flex items-center gap-1 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50">
              <Link2 size={12} /> Open in new tab
            </a>
            <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
          </div>
        </div>
        <div className="flex-1 overflow-hidden bg-slate-100">
          {kind === "pdf" && (
            <PdfViewer url={resource.url} title={resource.title} />
          )}
          {kind === "office" && (
            <OfficeViewer url={resource.url} title={resource.title} />
          )}
          {kind === "video" && (
            <video controls src={resource.url} className="h-full w-full bg-black" />
          )}
          {kind === "youtube" && youtubeUrl && (
            <iframe src={youtubeUrl} title={resource.title} className="h-full w-full border-0" allowFullScreen />
          )}
          {kind === "vimeo" && vimeoUrl && (
            <iframe src={vimeoUrl} title={resource.title} className="h-full w-full border-0" allowFullScreen />
          )}
          {(kind === "external" || ((kind === "youtube" && !youtubeUrl) || (kind === "vimeo" && !vimeoUrl))) && (
            <div className="flex h-full flex-col items-center justify-center gap-3 p-8 text-center text-slate-400">
              <FileText size={28} />
              <p className="text-sm">This resource can't be previewed inline — it'll open in a new tab instead.</p>
              <a href={resource.url} target="_blank" rel="noreferrer" className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
                <Link2 size={14} /> Open "{resource.title}"
              </a>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function ParticipantResourcesPage() {
  const { currentUser, projects, updateParticipant } = useAppData();
  const currentProject = projects.find((p) => p.id === (currentUser.projectId || "default")) || projects[0];
  const resources = currentProject?.resources || [];
  const completedIds = currentUser.completedResourceIds || [];
  const [viewingResource, setViewingResource] = useState(null);

  const toggleComplete = (resourceId) => {
    const has = completedIds.includes(resourceId);
    const next = has ? completedIds.filter((id) => id !== resourceId) : [...completedIds, resourceId];
    updateParticipant(currentUser.employeeId, { completedResourceIds: next });
  };

  const completedCount = resources.filter((r) => completedIds.includes(r.id)).length;

  return (
    <div className="mx-auto max-w-3xl space-y-6 px-6 py-6">
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <LogoWatermark size={200} className="-right-8 -bottom-12" />
        <div className="flex items-center gap-4">
          <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/15"><FileText size={22} /></div>
          <div>
            <h1 className="text-2xl font-bold">Resources</h1>
            <p className="text-sm text-white/70">{resources.length > 0 ? `${completedCount} of ${resources.length} completed` : "Documents and videos for this program"}</p>
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        {resources.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-14 text-center text-slate-400">
            <FileText size={26} className="mb-2" />
            <p className="text-sm">No resources have been added for your program yet.</p>
          </div>
        ) : (
          <div className="space-y-2">
            {resources.map((r) => {
              const done = completedIds.includes(r.id);
              return (
                <div key={r.id} className={`flex items-center justify-between gap-3 rounded-xl border px-4 py-3 ${done ? "border-emerald-100 bg-emerald-50" : "border-slate-100 bg-white"}`}>
                  <button onClick={() => setViewingResource(r)} className="flex flex-1 items-center gap-3 text-left">
                    <div className={`flex h-9 w-9 items-center justify-center rounded-lg ${done ? "bg-emerald-100 text-emerald-700" : "bg-blue-50 text-blue-700"}`}>
                      {r.type === "video" ? <FileType size={16} /> : <FileText size={16} />}
                    </div>
                    <div>
                      <div className="text-sm font-medium text-slate-800">{r.title}</div>
                      <div className="text-xs text-slate-400">{r.type === "video" ? "Video" : "Document"} · Click to view</div>
                    </div>
                  </button>
                  <button
                    onClick={() => toggleComplete(r.id)}
                    className={`flex shrink-0 items-center gap-1.5 rounded-lg px-3 py-2 text-xs font-semibold ${
                      done ? "bg-emerald-600 text-white hover:bg-emerald-700" : "border border-slate-200 text-slate-600 hover:bg-slate-50"
                    }`}
                  >
                    <CheckCircle2 size={13} /> {done ? "Completed" : "Mark Complete"}
                  </button>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {viewingResource && <ResourceViewerModal resource={viewingResource} onClose={() => setViewingResource(null)} />}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Participant Group — a dedicated view of the participant's own group and
// who else is in it, separate from the small teaser card on the Dashboard.
// ---------------------------------------------------------------------------

function ParticipantGroupPage() {
  const { currentUser, projects, participants } = useAppData();
  const myProject = projects.find((p) => p.id === (currentUser.projectId || "default")) || projects[0];
  const myGroup = (myProject?.groups || []).find((g) => g.participantIds.includes(currentUser.employeeId));
  const groupmates = myGroup ? participants.filter((p) => myGroup.participantIds.includes(p.employeeId) && p.employeeId !== currentUser.employeeId) : [];

  return (
    <div className="mx-auto max-w-3xl space-y-6 px-6 py-6">
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <LogoWatermark size={200} className="-right-8 -bottom-12" />
        <div className="flex items-center gap-4">
          <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/15"><Users size={22} /></div>
          <div>
            <h1 className="text-2xl font-bold">My Group</h1>
            <p className="text-sm text-white/70">{myGroup ? myGroup.name : "You haven't been assigned to a group yet"}</p>
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        {!myGroup ? (
          <div className="flex flex-col items-center justify-center py-14 text-center text-slate-400">
            <Users size={26} className="mb-2" />
            <p className="text-sm">You're not part of a group yet — check back once your program admin sets one up.</p>
          </div>
        ) : (
          <>
            <div className="mb-4 flex items-center gap-3 rounded-xl bg-blue-50 px-4 py-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-full bg-blue-800 text-sm font-semibold text-white">{initialsFromName(currentUser.name)}</div>
              <div>
                <div className="text-sm font-medium text-slate-800">{currentUser.name} <span className="text-xs font-normal text-blue-600">(you)</span></div>
                <div className="text-xs text-slate-400">{currentUser.cohort}</div>
              </div>
            </div>
            {groupmates.length === 0 ? (
              <p className="text-sm text-slate-400">No other participants in your group yet.</p>
            ) : (
              <div className="space-y-2">
                {groupmates.map((p) => (
                  <div key={p.employeeId} className="flex items-center gap-3 rounded-lg bg-slate-50 px-3 py-2.5">
                    <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-slate-400 text-xs font-semibold text-white">{initialsFromName(p.name)}</div>
                    <div className="min-w-0 flex-1">
                      <div className="text-sm font-medium text-slate-700">{p.name}</div>
                      <div className="text-xs text-slate-400">{p.cohort}{p.dept ? ` · ${p.dept}` : ""}</div>
                    </div>
                    <CopyEmailButton email={p.email} className="flex shrink-0 items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-blue-700 hover:bg-blue-50" />
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function ParticipantMeetingsPage() {
  const { currentUser } = useAppData();
  const { meetings, loading } = useSharedMeetings();
  const [qrSession, setQrSession] = useState(null);

  const today = new Date().toISOString().slice(0, 10);
  const upcoming = meetings
    .filter((m) => (m.date || "") >= today && (m.projectId || "default") === (currentUser.projectId || "default") && audienceMatches(m.audience, currentUser.cohort))
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));

  if (qrSession) {
    return <PresentAttendanceView session={qrSession} onExit={() => setQrSession(null)} />;
  }

  return (
    <div className="mx-auto max-w-3xl space-y-6 px-6 py-6">
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <LogoWatermark size={200} className="-right-8 -bottom-12" />
        <div className="flex items-center gap-4">
          <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/15"><Calendar size={22} /></div>
          <div>
            <h1 className="text-2xl font-bold">Meetings</h1>
            <p className="text-sm text-white/70">Your upcoming meetings — scan or show the QR code to check in</p>
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        {loading ? (
          <div className="flex items-center justify-center py-10 text-slate-400"><RefreshCw size={20} className="animate-spin" /></div>
        ) : upcoming.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-14 text-center text-slate-400">
            <Calendar size={26} className="mb-2" />
            <p className="text-sm">No upcoming meetings for you right now.</p>
          </div>
        ) : (
          <div className="space-y-3">
            {upcoming.map((m) => (
              <div key={m.id} className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-slate-100 bg-white p-4">
                <div className="flex items-center gap-3">
                  <ClientLogo logoUrl={m.logoUrl} size={40} radius={8} />
                  <div>
                    <div className="font-medium text-slate-800">{m.name}</div>
                    <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 text-xs text-slate-400">
                      <span className="flex items-center gap-1"><Calendar size={11} /> {m.date}{m.startTime ? ` · ${m.startTime}${m.endTime ? `–${m.endTime}` : ""}` : ""}</span>
                      {m.location && <span className="flex items-center gap-1"><MapPin size={11} /> {m.location}</span>}
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => setQrSession(m)}
                  className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-900"
                >
                  <QrCode size={14} /> Show QR Code
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Survey Analytics
// ---------------------------------------------------------------------------

const SURVEYS_STORAGE_KEY = "mercerassess-surveys";
const ATTENDANCE_STORAGE_KEY = "mercerassess-attendance-sessions";
const attendanceSessionsSeed = [];

function AddSurveyModal({ onClose, onCreate }) {
  const { participants, projects, currentProjectId } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId);
  const cohortOptions = ["All Participants", ...Array.from(new Set(participants.map((p) => p.cohort).filter(Boolean))).sort()];
  const [mode, setMode] = useState("upload"); // 'upload' | 'send' | 'build'
  const [name, setName] = useState("");
  const [audience, setAudience] = useState(["All Participants"]);
  const [message, setMessage] = useState("");
  const [fileName, setFileName] = useState("");
  const [expectedResponses, setExpectedResponses] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [parseError, setParseError] = useState(null);
  const [parsedAnalysis, setParsedAnalysis] = useState(null);
  const [builtQuestions, setBuiltQuestions] = useState([
    { key: 1, text: "", type: "rating" },
    { key: 2, text: "", type: "rating" },
  ]);
  const fileInputRef = React.useRef(null);

  const addBuiltQuestion = () => setBuiltQuestions((qs) => [...qs, { key: Date.now(), text: "", type: "rating" }]);
  const removeBuiltQuestion = (key) => setBuiltQuestions((qs) => qs.filter((q) => q.key !== key));
  const updateBuiltQuestion = (key, field, value) => setBuiltQuestions((qs) => qs.map((q) => (q.key === key ? { ...q, [field]: value } : q)));

  const audienceCount = (aud) => {
    const arr = Array.isArray(aud) ? aud : [aud];
    if (arr.includes("All Participants")) return participants.length;
    return participants.filter((p) => arr.includes(p.cohort)).length;
  };

  const handleFile = async (file) => {
    if (!file) return;
    setFileName(file.name);
    setParsing(true);
    setParseError(null);
    setParsedAnalysis(null);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const analysis = analyzeSurveyWorkbook(workbook);
      setParsedAnalysis(analysis);
    } catch (err) {
      setParseError(err.message || "Couldn't read that file. Try a .xlsx, .xls, or .csv export.");
    } finally {
      setParsing(false);
    }
  };

  const canSubmit =
    name.trim().length > 0 &&
    !parsing &&
    (mode === "send" || (mode === "upload" && fileName.trim().length > 0) || (mode === "build" && builtQuestions.some((q) => q.text.trim().length > 0)));

  const handleSubmit = () => {
    const today = new Date().toISOString().slice(0, 10);
    if (mode === "send") {
      onCreate({
        id: `survey-${Date.now()}`,
        name: name.trim(),
        responses: 0,
        rate: 0,
        uploaded: today,
        sourceType: "send",
        audience,
        logoUrl: currentProject?.logoUrl,
      });
    } else if (mode === "build") {
      const questions = builtQuestions
        .filter((q) => q.text.trim().length > 0)
        .map((q, i) => ({ id: `q${i + 1}`, text: q.text.trim(), type: q.type }));
      onCreate({
        id: `survey-${Date.now()}`,
        name: name.trim(),
        responses: 0,
        rate: 0,
        uploaded: today,
        sourceType: "built",
        questions,
        audience,
        logoUrl: currentProject?.logoUrl,
      });
    } else if (parsedAnalysis) {
      onCreate({
        id: `survey-${Date.now()}`,
        name: name.trim(),
        responses: parsedAnalysis.responses,
        rate: parsedAnalysis.rate,
        uploaded: today,
        sourceType: "upload",
        fileName: fileName.trim(),
        analysis: parsedAnalysis,
      });
    } else {
      // No file was successfully parsed (unsupported format, or skipped) —
      // fall back to a placeholder entry using generic mock charts.
      const responses = expectedResponses.trim() ? Number(expectedResponses) : Math.floor(40 + Math.random() * 120);
      onCreate({
        id: `survey-${Date.now()}`,
        name: name.trim(),
        responses,
        rate: Math.min(100, Math.round(40 + Math.random() * 55)),
        uploaded: today,
        sourceType: "upload",
        fileName: fileName.trim(),
      });
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Add Survey</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>

        <div className="mb-4 flex rounded-lg border border-slate-200 p-1">
          <button
            type="button"
            onClick={() => setMode("upload")}
            className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-2 text-sm font-medium ${mode === "upload" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-slate-50"}`}
          >
            <UploadCloud size={14} /> Upload File
          </button>
          <button
            type="button"
            onClick={() => setMode("send")}
            className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-2 text-sm font-medium ${mode === "send" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-slate-50"}`}
          >
            <Users size={14} /> Send to Participants
          </button>
          <button
            type="button"
            onClick={() => setMode("build")}
            className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-2 text-sm font-medium ${mode === "build" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-slate-50"}`}
          >
            <Pencil size={14} /> Build Survey
          </button>
        </div>

        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Survey name</label>
            <input
              autoFocus
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="e.g. Mid-Year Engagement Pulse"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
          </div>

          {mode === "upload" ? (
            <>
              <div>
                <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Survey results file</label>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleFile(e.target.files && e.target.files[0])}
                />
                <div
                  onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                  onDragLeave={() => setDragOver(false)}
                  onDrop={(e) => {
                    e.preventDefault();
                    setDragOver(false);
                    const file = e.dataTransfer.files && e.dataTransfer.files[0];
                    if (file) handleFile(file);
                  }}
                  onClick={() => fileInputRef.current && fileInputRef.current.click()}
                  className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-4 py-6 text-center transition-colors ${
                    dragOver ? "border-blue-400 bg-blue-50" : "border-slate-200 bg-slate-50"
                  }`}
                >
                  {parsing ? (
                    <>
                      <RefreshCw size={22} className="animate-spin text-blue-500" />
                      <p className="mt-2 text-sm font-medium text-slate-600">Reading {fileName}…</p>
                    </>
                  ) : (
                    <>
                      <UploadCloud size={22} className={dragOver ? "text-blue-500" : "text-slate-400"} />
                      <p className="mt-2 text-sm font-medium text-slate-600">
                        {fileName ? fileName : "Drag & drop an Excel or CSV file, or click to browse"}
                      </p>
                    </>
                  )}
                </div>
              </div>

              {parseError && (
                <div className="rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{parseError}</div>
              )}

              {parsedAnalysis && !parsing && (
                <div className="rounded-lg bg-emerald-50 px-3 py-2.5 text-xs text-emerald-700">
                  <div className="flex items-center gap-1.5 font-semibold"><CheckCircle2 size={13} /> Parsed successfully</div>
                  <div className="mt-1 text-emerald-600">
                    {parsedAnalysis.responses} response{parsedAnalysis.responses === 1 ? "" : "s"} detected across {parsedAnalysis.questionCount} question{parsedAnalysis.questionCount === 1 ? "" : "s"}.
                    {parsedAnalysis.keywords && " Open-text answers found and summarized."}
                  </div>
                </div>
              )}

              {!parsedAnalysis && !parsing && (
                <div>
                  <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Expected responses (optional)</label>
                  <input
                    type="number"
                    min="0"
                    value={expectedResponses}
                    onChange={(e) => setExpectedResponses(e.target.value)}
                    placeholder="e.g. 120"
                    className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
                  />
                  <p className="mt-1 text-xs text-slate-400">Used as a placeholder if the file above can't be auto-analyzed.</p>
                </div>
              )}
            </>
          ) : mode === "build" ? (
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Questions</label>
              <div className="space-y-2">
                {builtQuestions.map((q, i) => (
                  <div key={q.key} className="flex items-start gap-2">
                    <span className="mt-2.5 text-xs text-slate-400" style={{ width: 16 }}>{i + 1}.</span>
                    <input
                      value={q.text}
                      onChange={(e) => updateBuiltQuestion(q.key, "text", e.target.value)}
                      placeholder="e.g. How clear were the session objectives?"
                      className="flex-1 rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
                    />
                    <select
                      value={q.type}
                      onChange={(e) => updateBuiltQuestion(q.key, "type", e.target.value)}
                      className="rounded-lg border border-slate-200 px-2 py-2 text-xs"
                    >
                      <option value="rating">Rating 1–5</option>
                      <option value="text">Open text</option>
                    </select>
                    <button
                      type="button"
                      onClick={() => removeBuiltQuestion(q.key)}
                      className="mt-1.5 text-slate-300 hover:text-red-500"
                      disabled={builtQuestions.length <= 1}
                    >
                      <X size={15} />
                    </button>
                  </div>
                ))}
              </div>
              <button
                type="button"
                onClick={addBuiltQuestion}
                className="mt-2 flex items-center gap-1.5 text-sm font-medium text-blue-700 hover:underline"
              >
                <Plus size={14} /> Add question
              </button>
              <p className="mt-3 rounded-lg bg-blue-50 px-3 py-2.5 text-xs text-blue-700">
                After creating this survey, you'll get a shareable link and QR code participants can use to answer it — responses collect automatically as people submit.
              </p>
              <label className="mb-1 mt-4 block text-xs font-semibold uppercase text-slate-400">Audience</label>
              <AudienceSelector value={audience} onChange={setAudience} cohortOptions={cohortOptions} />
              <p className="mt-1 text-xs text-slate-400">{audienceCount(audience)} participant{audienceCount(audience) === 1 ? "" : "s"} will see this in their Surveys tab</p>
            </div>
          ) : (
            <>
              <div>
                <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Send to</label>
                <AudienceSelector value={audience} onChange={setAudience} cohortOptions={cohortOptions} />
                <p className="mt-1 text-xs text-slate-400">{audienceCount(audience)} participant{audienceCount(audience) === 1 ? "" : "s"} will receive this survey</p>
              </div>
              <div>
                <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Message (optional)</label>
                <textarea
                  rows={3}
                  value={message}
                  onChange={(e) => setMessage(e.target.value)}
                  placeholder="A short note to include with the survey invite…"
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
                />
              </div>
            </>
          )}
        </div>

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Cancel
          </button>
          <button
            disabled={!canSubmit}
            onClick={handleSubmit}
            className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {mode === "send" ? <><Users size={14} /> Send Survey</> : mode === "build" ? <><Pencil size={14} /> Create Survey</> : <><UploadCloud size={14} /> Add Survey</>}
          </button>
        </div>
      </div>
    </div>
  );
}

// Chart bars only have room for a short, truncated question label, so this
// custom tooltip shows the full text (falling back to the short label for
// generic mock surveys that don't have a separate full version).
function SurveyQuestionTooltip({ active, payload, mode }) {
  if (!active || !payload || !payload.length) return null;
  const row = payload[0].payload;
  return (
    <div className="rounded-lg border border-slate-200 bg-white p-3 shadow-lg" style={{ maxWidth: 280 }}>
      <div className="mb-1.5 text-xs font-semibold text-slate-700">{row.fullQuestion || row.question}</div>
      {mode === "avg" ? (
        <div className="text-sm font-semibold text-teal-600">{row.avg} / 5</div>
      ) : (
        <div className="space-y-1">
          {LIKERT_LABELS_5.map((label, i) => (
            <div key={label} className="flex items-center justify-between gap-3 text-xs">
              <span className="flex items-center gap-1.5 text-slate-500">
                <span className="h-2 w-2 rounded-full" style={{ backgroundColor: LIKERT_COLORS_5[i] }} />
                {label}
              </span>
              <span className="font-medium text-slate-700">{row[label]}%</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Share modal for admin-built surveys — a copyable link plus a QR code
// (rendered via a public QR image API, so no extra dependency is needed)
// that opens the no-login TakeSurveyPage for that survey.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Attendance — separate from surveys entirely. A session is just a name +
// date; people scan its QR code, type their name, and tap Check In. No
// questions, no rating, just a live headcount and roster.
// ---------------------------------------------------------------------------

function AddAttendanceSessionModal({ onClose, onSave, initial, onShare, onPresent, onRemind, onDeleteRequest }) {
  const { participants, projects, currentProjectId } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId);
  const cohortOptions = ["All Participants", ...Array.from(new Set(participants.map((p) => p.cohort).filter(Boolean))).sort()];
  const isEdit = Boolean(initial);
  const [name, setName] = useState(initial?.name || "");
  const [date, setDate] = useState(initial?.date || new Date().toISOString().slice(0, 10));
  const [startTime, setStartTime] = useState(initial?.startTime || "10:00");
  const [endTime, setEndTime] = useState(initial?.endTime || "11:00");
  const [location, setLocation] = useState(initial?.location || "");
  const [description, setDescription] = useState(initial?.description || "");
  const [audience, setAudience] = useState(() => {
    const initialAudience = initial?.audience;
    if (!initialAudience) return ["All Participants"];
    return Array.isArray(initialAudience) ? initialAudience : [initialAudience];
  });

  const audienceCount = (aud) => {
    const arr = Array.isArray(aud) ? aud : [aud];
    if (arr.includes("All Participants")) return participants.length;
    return participants.filter((p) => arr.includes(p.cohort)).length;
  };

  const handleSubmit = () => {
    onSave({
      id: initial?.id || `attendance-${Date.now()}`,
      name: name.trim(),
      date,
      startTime,
      endTime,
      location: location.trim(),
      description: description.trim(),
      audience,
      logoUrl: initial?.logoUrl ?? currentProject?.logoUrl,
      projectId: initial?.projectId ?? currentProjectId,
      createdAt: initial?.createdAt || new Date().toISOString(),
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">{isEdit ? "Edit Meeting" : "New Meeting"}</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>
        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Meeting title</label>
            <input
              autoFocus
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="e.g. Leadership Workshop — Day 1"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
          </div>
          <div className="grid grid-cols-3 gap-3">
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Date</label>
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Start</label>
              <input
                type="time"
                value={startTime}
                onChange={(e) => setStartTime(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">End</label>
              <input
                type="time"
                value={endTime}
                onChange={(e) => setEndTime(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
              />
            </div>
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Location <span className="font-normal normal-case text-slate-400">(optional)</span></label>
            <input
              value={location}
              onChange={(e) => setLocation(e.target.value)}
              placeholder="e.g. Training Room B, or a video call link"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Description <span className="font-normal normal-case text-slate-400">(optional)</span></label>
            <textarea
              rows={2}
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="What this meeting covers…"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Audience</label>
            <AudienceSelector value={audience} onChange={setAudience} cohortOptions={cohortOptions} />
            <p className="mt-1 text-xs text-slate-400">{audienceCount(audience)} participant{audienceCount(audience) === 1 ? "" : "s"} in this audience</p>
          </div>
        </div>

        {isEdit && (
          <div className="mt-4 flex flex-wrap gap-2 border-t border-slate-100 pt-4">
            <button onClick={onShare} className="flex items-center gap-1.5 rounded-lg bg-blue-50 px-3 py-1.5 text-xs font-medium text-blue-700 hover:bg-blue-100">
              <QrCode size={12} /> Share
            </button>
            <button onClick={onPresent} className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50">
              <Maximize2 size={12} /> Present
            </button>
            <button onClick={onRemind} className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50">
              <Mail size={12} /> Remind
            </button>
            <button onClick={onDeleteRequest} className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-red-500 hover:bg-red-50">
              <Trash2 size={12} /> Delete
            </button>
          </div>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Cancel
          </button>
          <button
            disabled={!name.trim()}
            onClick={handleSubmit}
            className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
          >
            <CalendarPlus size={14} /> {isEdit ? "Save Changes" : "Create Meeting"}
          </button>
        </div>
      </div>
    </div>
  );
}

function ShareAttendanceModal({ session, onClose, onPreview }) {
  const [copied, setCopied] = useState(false);
  const url = `${window.location.origin}${window.location.pathname}?attendance=${session.id}`;

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(url);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      // clipboard API can fail without permission — link is still shown to copy manually
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-sm rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Check-in QR — "{session.name}"</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>
        <div className="flex flex-col items-center rounded-xl border border-slate-100 bg-slate-50 p-5">
          <QRCodeImage value={url} size={200} />
          <p className="mt-3 text-center text-xs text-slate-500">Scan to check in — no login required.</p>
        </div>
        <div className="mt-4">
          <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Shareable link</label>
          <div className="flex items-center gap-2">
            <input readOnly value={url} className="flex-1 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-600" />
            <button onClick={copyLink} className="flex items-center gap-1 rounded-lg bg-blue-800 px-3 py-2 text-xs font-semibold text-white hover:bg-blue-900">
              {copied ? <CheckCircle2 size={13} /> : <Copy size={13} />} {copied ? "Copied" : "Copy"}
            </button>
          </div>
        </div>
        {onPreview && (
          <button
            onClick={onPreview}
            className="mt-3 flex w-full items-center justify-center gap-1.5 rounded-lg bg-blue-50 px-4 py-2.5 text-sm font-semibold text-blue-700 hover:bg-blue-100"
          >
            <QrCode size={14} /> Preview check-in form
          </button>
        )}
        <div className="mt-4 flex justify-end">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Done
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Week calendar grid — meetings positioned by day/time, click to edit.
// ---------------------------------------------------------------------------

function startOfWeek(date) {
  const d = new Date(date);
  d.setDate(d.getDate() - d.getDay());
  d.setHours(0, 0, 0, 0);
  return d;
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

function dateKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function timeToDecimal(t) {
  if (!t) return null;
  const [h, m] = t.split(":").map(Number);
  return h + (m || 0) / 60;
}

const CAL_HOUR_START = 7;
const CAL_HOUR_END = 20;
const CAL_ROW_HEIGHT = 52;

function CalendarWeekView({ sessions, weekStart, onPrevWeek, onNextWeek, onToday, onSelectMeeting }) {
  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));
  const totalHeight = (CAL_HOUR_END - CAL_HOUR_START) * CAL_ROW_HEIGHT;
  const todayKey = dateKey(new Date());
  const meetingsByDay = days.map((day) => sessions.filter((s) => s.date === dateKey(day)));

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <button onClick={onToday} className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50">Today</button>
          <button onClick={onPrevWeek} className="rounded-lg border border-slate-200 p-1.5 text-slate-500 hover:bg-slate-50"><ChevronRight size={14} style={{ transform: "scaleX(-1)" }} /></button>
          <button onClick={onNextWeek} className="rounded-lg border border-slate-200 p-1.5 text-slate-500 hover:bg-slate-50"><ChevronRight size={14} /></button>
          <span className="ml-2 text-sm font-semibold text-slate-700">
            {days[0].toLocaleDateString(undefined, { month: "short", day: "numeric" })} – {days[6].toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" })}
          </span>
        </div>
      </div>

      <div style={{ overflowX: "auto" }}>
        <div style={{ minWidth: 720 }}>
          <div className="grid" style={{ gridTemplateColumns: "44px repeat(7, 1fr)" }}>
            <div />
            {days.map((day, i) => {
              const isToday = dateKey(day) === todayKey;
              return (
                <div key={i} className="border-b border-slate-100 pb-2 text-center">
                  <div className="text-xs text-slate-400">{day.toLocaleDateString(undefined, { weekday: "short" })}</div>
                  <div
                    className="mx-auto mt-0.5 flex h-7 w-7 items-center justify-center rounded-full text-sm font-semibold"
                    style={isToday ? { backgroundColor: BRAND_NAVY, color: "white" } : { color: "#334155" }}
                  >
                    {day.getDate()}
                  </div>
                </div>
              );
            })}
          </div>

          <div className="relative grid" style={{ gridTemplateColumns: "44px repeat(7, 1fr)", height: totalHeight }}>
            <div className="relative">
              {Array.from({ length: CAL_HOUR_END - CAL_HOUR_START }, (_, i) => {
                const h = CAL_HOUR_START + i;
                const ampm = h >= 12 ? "PM" : "AM";
                const hh = h % 12 === 0 ? 12 : h % 12;
                return (
                  <div key={i} className="absolute w-full pr-2 text-right text-slate-400" style={{ top: i * CAL_ROW_HEIGHT - 6, fontSize: 10 }}>
                    {hh} {ampm}
                  </div>
                );
              })}
            </div>
            {days.map((day, dayIdx) => (
              <div key={dayIdx} className="relative border-l border-slate-200">
                {Array.from({ length: CAL_HOUR_END - CAL_HOUR_START }, (_, i) => (
                  <div key={`h${i}`} className="absolute w-full border-t border-slate-200" style={{ top: i * CAL_ROW_HEIGHT }} />
                ))}
                {Array.from({ length: CAL_HOUR_END - CAL_HOUR_START }, (_, i) => (
                  <div key={`hh${i}`} className="absolute w-full border-t border-dashed border-slate-100" style={{ top: i * CAL_ROW_HEIGHT + CAL_ROW_HEIGHT / 2 }} />
                ))}
                <div className="absolute w-full border-t border-slate-200" style={{ top: totalHeight }} />
                {meetingsByDay[dayIdx].map((s) => {
                  const start = timeToDecimal(s.startTime) ?? 9;
                  const end = timeToDecimal(s.endTime) ?? start + 1;
                  const top = Math.max(0, (start - CAL_HOUR_START) * CAL_ROW_HEIGHT);
                  const height = Math.max(26, (end - start) * CAL_ROW_HEIGHT);
                  return (
                    <button
                      key={s.id}
                      onClick={() => onSelectMeeting(s)}
                      className="absolute overflow-hidden rounded-md px-1.5 py-1 text-left hover:opacity-80"
                      style={{ top, height, left: 2, right: 2, backgroundColor: "#dbeafe", color: BRAND_NAVY }}
                    >
                      <div className="truncate font-semibold" style={{ fontSize: 11 }}>{s.name}</div>
                      <div className="truncate text-blue-600" style={{ fontSize: 10 }}>
                        {s.startTime}{s.endTime ? `–${s.endTime}` : ""}
                      </div>
                    </button>
                  );
                })}
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Attendance Report — aggregates check-ins across every meeting into a
// per-participant attendance rate, rather than one meeting at a time.
// ---------------------------------------------------------------------------

function AttendanceReportPanel() {
  const { participants, currentProjectId } = useAppData();
  const projectParticipants = participants.filter((p) => (p.projectId || "default") === currentProjectId);
  const [rows, setRows] = useState(null);
  const [loading, setLoading] = useState(true);
  const [sortBy, setSortBy] = useState("rate"); // 'rate' | 'name'

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      try {
        const res = await fetch(`/api/data?key=${encodeURIComponent(ATTENDANCE_STORAGE_KEY)}`);
        const data = res.ok ? await res.json() : null;
        const allMeetings = data ? JSON.parse(data.value || "[]") : [];
        const meetings = allMeetings.filter((m) => (m.projectId || "default") === currentProjectId);
        if (cancelled) return;

        const checkinsByMeeting = {};
        for (const m of meetings) {
          try {
            const cres = await fetch(`/api/attendance?sessionId=${encodeURIComponent(m.id)}`);
            const cdata = cres.ok ? await cres.json() : { checkins: [] };
            checkinsByMeeting[m.id] = cdata.checkins || [];
          } catch (err) {
            checkinsByMeeting[m.id] = [];
          }
        }
        if (cancelled) return;

        const computed = projectParticipants.map((p) => {
          const applicable = meetings.filter((m) => audienceMatches(m.audience, p.cohort));
          const attendedMeetings = applicable.filter((m) => (checkinsByMeeting[m.id] || []).some((c) => c.employeeId === p.employeeId));
          const lastEntries = attendedMeetings
            .map((m) => (checkinsByMeeting[m.id] || []).find((c) => c.employeeId === p.employeeId))
            .filter(Boolean)
            .sort((a, b) => new Date(b.checkedInAt) - new Date(a.checkedInAt));
          return {
            employeeId: p.employeeId,
            name: p.name,
            cohort: p.cohort,
            applicable: applicable.length,
            attended: attendedMeetings.length,
            rate: applicable.length > 0 ? Math.round((attendedMeetings.length / applicable.length) * 100) : null,
            lastAttended: lastEntries[0] ? lastEntries[0].checkedInAt : null,
          };
        });
        if (!cancelled) setRows(computed);
      } catch (err) {
        if (!cancelled) setRows([]);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projectParticipants.length, currentProjectId]);

  const sorted = rows
    ? [...rows].sort((a, b) => {
        if (sortBy === "name") return a.name.localeCompare(b.name);
        return (a.rate ?? 999) - (b.rate ?? 999);
      })
    : [];

  const withData = sorted.filter((r) => r.rate !== null);
  const avgRate = withData.length ? Math.round(withData.reduce((s, r) => s + r.rate, 0) / withData.length) : 0;
  const atRiskCount = withData.filter((r) => r.rate < 50).length;

  const rateColor = (rate) => {
    if (rate === null) return { bg: "#f1f5f9", text: "#94a3b8" };
    if (rate >= 80) return { bg: "#d1fae5", text: "#047857" };
    if (rate >= 50) return { bg: "#fef3c7", text: "#b45309" };
    return { bg: "#fee2e2", text: "#dc2626" };
  };

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-bold text-slate-800">Attendance Report</h2>
          <p className="text-sm text-slate-400">Attendance rate per participant, aggregated across every meeting that applies to them</p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={() => setSortBy("rate")}
            className={`rounded-lg px-3 py-1.5 text-xs font-medium ${sortBy === "rate" ? "bg-blue-800 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}
          >
            Sort: Lowest rate first
          </button>
          <button
            onClick={() => setSortBy("name")}
            className={`rounded-lg px-3 py-1.5 text-xs font-medium ${sortBy === "name" ? "bg-blue-800 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}
          >
            Sort: Name
          </button>
        </div>
      </div>

      {loading ? (
        <div className="flex items-center justify-center py-14 text-slate-400"><RefreshCw size={22} className="animate-spin" /></div>
      ) : rows.length === 0 ? (
        <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 py-14 text-center text-slate-400">
          <UserCheck size={26} className="mb-2" />
          <p className="text-sm">No participants to report on yet.</p>
        </div>
      ) : (
        <>
          <div className="mb-4 grid grid-cols-2 gap-4 sm:grid-cols-3">
            <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
              <div className="text-xs uppercase text-slate-400">Average Attendance</div>
              <div className="text-2xl font-bold text-slate-800">{avgRate}%</div>
            </div>
            <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
              <div className="text-xs uppercase text-slate-400">At Risk (&lt;50%)</div>
              <div className="text-2xl font-bold text-red-500">{atRiskCount}</div>
            </div>
            <div className="rounded-xl border border-slate-100 bg-slate-50 p-4">
              <div className="text-xs uppercase text-slate-400">Participants Tracked</div>
              <div className="text-2xl font-bold text-slate-800">{rows.length}</div>
            </div>
          </div>

          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
                <th className="py-2 font-semibold">Participant</th>
                <th className="py-2 font-semibold">Cohort</th>
                <th className="py-2 text-center font-semibold">Attended</th>
                <th className="py-2 font-semibold">Rate</th>
                <th className="py-2 font-semibold">Last Attended</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((r) => {
                const colors = rateColor(r.rate);
                return (
                  <tr key={r.employeeId} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                    <td className="py-2.5">
                      <div className="flex items-center gap-2.5">
                        <div className="flex h-7 w-7 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">{initialsFromName(r.name)}</div>
                        <span className="font-medium text-slate-800">{r.name}</span>
                      </div>
                    </td>
                    <td className="py-2.5"><span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-medium text-blue-700">{r.cohort}</span></td>
                    <td className="py-2.5 text-center text-slate-600">{r.attended} / {r.applicable}</td>
                    <td className="py-2.5">
                      {r.rate === null ? (
                        <span className="text-xs text-slate-400">No meetings yet</span>
                      ) : (
                        <div className="flex items-center gap-2">
                          <div className="h-2 rounded-full bg-slate-100" style={{ width: 90 }}>
                            <div className="h-2 rounded-full" style={{ width: `${r.rate}%`, backgroundColor: colors.text }} />
                          </div>
                          <span className="rounded-full px-2 py-0.5 text-xs font-semibold" style={{ backgroundColor: colors.bg, color: colors.text }}>{r.rate}%</span>
                        </div>
                      )}
                    </td>
                    <td className="py-2.5 text-slate-500">{r.lastAttended ? new Date(r.lastAttended).toLocaleDateString() : "—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </>
      )}
    </div>
  );
}

function AttendancePanel({ onPreviewAttendance }) {
  const { showToast, currentProjectId } = useAppData();
  const [sessions, setSessions] = useState(attendanceSessionsSeed);
  const [loaded, setLoaded] = useState(false);
  const [modal, setModal] = useState(null); // null | { mode: "add" } | { mode: "edit", data }
  const [weekStart, setWeekStart] = useState(startOfWeek(new Date()));
  const [shareSession, setShareSession] = useState(null);
  const [presentSession, setPresentSession] = useState(null);
  const [reminderSession, setReminderSession] = useState(null);
  const [expandedId, setExpandedId] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(`/api/data?key=${encodeURIComponent(ATTENDANCE_STORAGE_KEY)}`);
        if (res.ok) {
          const data = await res.json();
          const parsed = JSON.parse(data.value || "[]");
          if (!cancelled && Array.isArray(parsed)) setSessions(parsed);
        }
      } catch (err) {
        // nothing saved yet
      } finally {
        if (!cancelled) setLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!loaded) return;
    (async () => {
      try {
        await fetch(`/api/data?key=${encodeURIComponent(ATTENDANCE_STORAGE_KEY)}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ value: JSON.stringify(sessions) }),
        });
      } catch (err) {
        console.error("Failed to save attendance sessions", err);
        showToast("Couldn't save — check the database connection (see README)");
      }
    })();
  }, [sessions, loaded]);

  const saveSession = (session) => {
    setSessions((prev) => {
      const exists = prev.some((s) => s.id === session.id);
      return exists ? prev.map((s) => (s.id === session.id ? session : s)) : [session, ...prev];
    });
    setModal(null);
    showToast(modal?.mode === "edit" ? `Saved changes to "${session.name}"` : `Created "${session.name}"`);
  };

  const deleteSession = (id, name) => {
    setSessions((prev) => prev.filter((s) => s.id !== id));
    showToast(`Deleted "${name}"`);
    setConfirmDelete(null);
    setModal(null);
  };

  const today = new Date().toISOString().slice(0, 10);
  const projectSessions = sessions.filter((s) => (s.projectId || "default") === currentProjectId);
  const sorted = [...projectSessions].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const upcoming = sorted.filter((s) => (s.date || "") >= today);
  const past = sorted.filter((s) => (s.date || "") < today).reverse();

  const renderGroup = (title, list) => (
    <div className="mb-6 last:mb-0">
      <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">{title} <span className="font-normal normal-case">({list.length})</span></h3>
      <div className="space-y-3">
        {list.map((session) => (
          <AttendanceSessionRow
            key={session.id}
            session={session}
            expanded={expandedId === session.id}
            onToggleExpand={() => setExpandedId(expandedId === session.id ? null : session.id)}
            onShare={() => setShareSession(session)}
            onPresent={() => setPresentSession(session)}
            onSendReminder={() => setReminderSession(session)}
            onDelete={() => setConfirmDelete(session)}
          />
        ))}
      </div>
    </div>
  );

  return (
    <div>
      <div className="mb-4 flex items-center justify-between">
        <div>
          <h2 className="text-lg font-bold text-slate-800">Meetings</h2>
          <p className="text-sm text-slate-400">Click a meeting on the calendar to reschedule or manage it</p>
        </div>
        <button onClick={() => setModal({ mode: "add" })} className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
          <CalendarPlus size={15} /> New Meeting
        </button>
      </div>

      <div className="mb-6">
        <CalendarWeekView
          sessions={projectSessions}
          weekStart={weekStart}
          onToday={() => setWeekStart(startOfWeek(new Date()))}
          onPrevWeek={() => setWeekStart((w) => addDays(w, -7))}
          onNextWeek={() => setWeekStart((w) => addDays(w, 7))}
          onSelectMeeting={(session) => setModal({ mode: "edit", data: session })}
        />
      </div>

      {projectSessions.length === 0 ? (
        <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 py-14 text-center text-slate-400">
          <Calendar size={26} className="mb-2" />
          <p className="text-sm">No meetings scheduled yet. Create one to generate a check-in QR code.</p>
        </div>
      ) : (
        <>
          {upcoming.length > 0 && renderGroup("Upcoming", upcoming)}
          {past.length > 0 && renderGroup("Past", past)}
        </>
      )}

      {modal && (
        <AddAttendanceSessionModal
          onClose={() => setModal(null)}
          onSave={saveSession}
          initial={modal.mode === "edit" ? modal.data : null}
          onShare={() => { setShareSession(modal.data); setModal(null); }}
          onPresent={() => { setPresentSession(modal.data); setModal(null); }}
          onRemind={() => { setReminderSession(modal.data); setModal(null); }}
          onDeleteRequest={() => { setConfirmDelete(modal.data); setModal(null); }}
        />
      )}
      {shareSession && (
        <ShareAttendanceModal
          session={shareSession}
          onClose={() => setShareSession(null)}
          onPreview={() => { setShareSession(null); onPreviewAttendance && onPreviewAttendance(shareSession.id); }}
        />
      )}
      {presentSession && <PresentAttendanceView session={presentSession} onExit={() => setPresentSession(null)} />}
      {reminderSession && <SendReminderModal session={reminderSession} onClose={() => setReminderSession(null)} />}
      {confirmDelete && (
        <ConfirmModal
          title="Delete this meeting?"
          message={`"${confirmDelete.name}" and all of its check-in records will be permanently removed.`}
          onCancel={() => setConfirmDelete(null)}
          onConfirm={() => deleteSession(confirmDelete.id, confirmDelete.name)}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Send Reminders — composes and sends a real email via the /api/send-reminder
// route (Resend), to the personal email addresses of participants in the
// meeting's audience.
// ---------------------------------------------------------------------------

function SendReminderModal({ session, onClose }) {
  const { participants, showToast } = useAppData();
  const audienceParticipants = participants.filter((p) => audienceMatches(session.audience, p.cohort));

  const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const recipients = audienceParticipants.filter((p) => p.email && EMAIL_RE.test(p.email.trim()));
  const skippedCount = audienceParticipants.length - recipients.length;

  const defaultSubject = `Reminder: ${session.name}`;
  const defaultMessage = `Hi,\n\nThis is a reminder about "${session.name}"${session.date ? ` on ${session.date}` : ""}${session.startTime ? ` at ${session.startTime}` : ""}.${session.location ? `\n\nLocation: ${session.location}` : ""}${session.description ? `\n\n${session.description}` : ""}\n\nSee you there!`;

  const [subject, setSubject] = useState(defaultSubject);
  const [message, setMessage] = useState(defaultMessage);
  const [sending, setSending] = useState(false);
  const [result, setResult] = useState(null); // { ok, error }

  const handleSend = async () => {
    setSending(true);
    setResult(null);
    try {
      const res = await fetch("/api/send-reminder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          recipients: recipients.map((p) => p.email.trim()),
          subject,
          message,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        setResult({ ok: false, error: data.error || "Failed to send reminders." });
      } else {
        setResult({ ok: true });
        showToast(`Sent reminder to ${recipients.length} participant${recipients.length === 1 ? "" : "s"}`);
      }
    } catch (err) {
      setResult({ ok: false, error: "Network error while sending reminders." });
    } finally {
      setSending(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Send Reminders</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>

        <div className="mb-4 flex items-center gap-2 rounded-lg bg-blue-50 px-3 py-2.5 text-sm text-blue-700">
          <Mail size={15} /> Sending to {recipients.length} participant{recipients.length === 1 ? "" : "s"} in "{formatAudience(session.audience)}", using their personal email on file.
        </div>
        {skippedCount > 0 && (
          <div className="mb-4 flex items-center gap-2 rounded-lg bg-amber-50 px-3 py-2.5 text-xs text-amber-700">
            <XCircle size={13} /> Skipping {skippedCount} participant{skippedCount === 1 ? "" : "s"} with a missing or invalid email address.
          </div>
        )}

        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Subject</label>
            <input
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Message</label>
            <textarea
              rows={6}
              value={message}
              onChange={(e) => setMessage(e.target.value)}
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
          </div>
        </div>

        {result && !result.ok && (
          <div className="mt-3 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">
            {result.error}
            {String(result.error || "").includes("RESEND_API_KEY") && (
              <div className="mt-1">Make sure <code>RESEND_API_KEY</code> is set in this project's Vercel environment variables, then redeploy.</div>
            )}
          </div>
        )}
        {result && result.ok && (
          <div className="mt-3 flex items-center gap-1.5 rounded-lg bg-emerald-50 px-3 py-2.5 text-xs text-emerald-700">
            <CheckCircle2 size={13} /> Reminders sent successfully.
          </div>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Close
          </button>
          <button
            disabled={sending || recipients.length === 0}
            onClick={handleSend}
            className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {sending ? <RefreshCw size={14} className="animate-spin" /> : <Mail size={14} />}
            {sending ? "Sending…" : "Send Reminders"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Present mode — a big, high-contrast full-screen QR display meant to be
// projected on a screen for a room to scan, separate from the compact QR
// shown in the admin's own Share modal.
// ---------------------------------------------------------------------------

function PresentAttendanceView({ session, onExit }) {
  const url = `${window.location.origin}${window.location.pathname}?attendance=${session.id}`;
  return (
    <div className="fixed inset-0 z-50 flex flex-col items-center justify-center p-8 text-white" style={brandGradient()}>
      <LogoWatermark size={320} className="-right-16" style={{ bottom: -100 }} />
      <button onClick={onExit} className="absolute right-6 top-6 flex items-center gap-1.5 rounded-lg bg-white/15 px-4 py-2 text-sm font-medium hover:bg-white/25">
        <X size={15} /> Exit
      </button>
      <div className="relative flex flex-col items-center">
        {session.logoUrl && (
          <div className="mb-4 rounded-2xl bg-white p-3">
            <ClientLogo logoUrl={session.logoUrl} size={64} radius={10} />
          </div>
        )}
        <div className="mb-2 text-sm uppercase tracking-widest text-white/70">Scan to check in</div>
        <h1 className="mb-1 text-3xl font-bold">{session.name}</h1>
        <p className="mb-8 text-white/70">{session.date}{session.startTime ? ` · ${session.startTime}` : ""}{session.location ? ` · ${session.location}` : ""}</p>
        <div className="rounded-3xl bg-white p-6">
          <QRCodeImage value={url} size={340} />
        </div>
      </div>
    </div>
  );
}

function AttendanceSessionRow({ session, expanded, onToggleExpand, onShare, onDelete, onPresent, onSendReminder }) {
  const { checkins, loading, refresh } = useAttendanceCheckins(expanded ? session.id : null);

  return (
    <div className="overflow-hidden rounded-xl border border-slate-100">
      <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-50 px-4 py-3">
        <div>
          <div className="font-medium text-slate-800">{session.name}</div>
          <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 text-xs text-slate-400">
            <span className="flex items-center gap-1"><Calendar size={11} /> {session.date}{session.startTime ? ` · ${session.startTime}${session.endTime ? `–${session.endTime}` : ""}` : ""}</span>
            {session.location && <span className="flex items-center gap-1"><MapPin size={11} /> {session.location}</span>}
            <span className="rounded-full bg-blue-50 px-2 py-0.5 font-medium text-blue-700">{formatAudience(session.audience)}</span>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button onClick={onToggleExpand} className="flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50">
            <UserCheck size={13} /> {expanded ? "Hide" : "View"} check-ins
          </button>
          <button onClick={onSendReminder} className="flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50">
            <Mail size={13} /> Remind
          </button>
          <button onClick={onPresent} className="flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50">
            <Maximize2 size={13} /> Present
          </button>
          <button onClick={onShare} className="flex items-center gap-1.5 rounded-lg bg-blue-50 px-3 py-1.5 text-sm font-medium text-blue-700 hover:bg-blue-100">
            <QrCode size={13} /> Share
          </button>
          <button onClick={onDelete} className="text-slate-300 hover:text-red-500"><Trash2 size={15} /></button>
        </div>
      </div>
      {session.description && (
        <div className="border-t border-slate-100 px-4 py-2 text-sm text-slate-500">{session.description}</div>
      )}
      {expanded && (
        <div className="px-4 py-3">
          <div className="mb-2 flex items-center justify-between">
            <span className="text-sm font-semibold text-slate-700">{checkins.length} checked in</span>
            <button onClick={refresh} className="flex items-center gap-1 text-xs font-medium text-blue-700 hover:underline">
              <RefreshCw size={12} className={loading ? "animate-spin" : ""} /> Refresh
            </button>
          </div>
          {checkins.length === 0 ? (
            <p className="py-4 text-center text-sm text-slate-400">No check-ins yet.</p>
          ) : (
            <div style={{ maxHeight: 260, overflowY: "auto" }}>
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
                    <th className="py-1.5 font-semibold">Name</th>
                    <th className="py-1.5 font-semibold">Employee ID</th>
                    <th className="py-1.5 font-semibold">Cohort</th>
                    <th className="py-1.5 font-semibold">Checked in</th>
                  </tr>
                </thead>
                <tbody>
                  {checkins.map((c, i) => (
                    <tr key={i} className="border-b border-slate-50 last:border-0">
                      <td className="py-1.5 font-medium text-slate-700">{c.name}</td>
                      <td className="py-1.5 text-slate-500">
                        {c.employeeId || <span className="rounded-full bg-amber-50 px-2 py-0.5 text-xs font-medium text-amber-600">Unmatched</span>}
                      </td>
                      <td className="py-1.5 text-slate-500">{c.cohort ? `${c.cohort}${c.dept ? " · " + c.dept : ""}` : "—"}</td>
                      <td className="py-1.5 text-slate-400">{new Date(c.checkedInAt).toLocaleString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// Renders a QR code entirely client-side (no network call to any third-party
// image service — earlier versions of this app used one and it was
// unreliable in some environments). Generates a PNG data URI locally.
function QRCodeImage({ value, size = 200 }) {
  const [dataUri, setDataUri] = useState(null);
  const [error, setError] = useState(false);

  useEffect(() => {
    let cancelled = false;
    QRCode.toDataURL(value, { width: size * 2, margin: 1 })
      .then((uri) => { if (!cancelled) setDataUri(uri); })
      .catch(() => { if (!cancelled) setError(true); });
    return () => { cancelled = true; };
  }, [value, size]);

  if (error) {
    return (
      <div
        className="flex items-center justify-center rounded-lg bg-slate-100 text-center text-xs text-slate-400"
        style={{ width: size, height: size }}
      >
        Couldn't generate QR code
      </div>
    );
  }
  if (!dataUri) {
    return (
      <div
        className="flex items-center justify-center rounded-lg bg-slate-100"
        style={{ width: size, height: size }}
      >
        <RefreshCw size={18} className="animate-spin text-slate-300" />
      </div>
    );
  }
  return <img src={dataUri} alt="QR code" width={size} height={size} className="rounded-lg bg-white p-2" />;
}

// ---------------------------------------------------------------------------
// Send to Participants — notifies the survey's audience by real email
// (via the same /api/send-reminder route as meeting reminders) that a new
// survey is ready for them to answer in the in-app Surveys section.
// ---------------------------------------------------------------------------

function SendSurveyModal({ survey, onClose }) {
  const { participants, showToast } = useAppData();
  const audienceParticipants = participants.filter((p) => audienceMatches(survey.audience, p.cohort));

  const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const recipients = audienceParticipants.filter((p) => p.email && EMAIL_RE.test(p.email.trim()));
  const skippedCount = audienceParticipants.length - recipients.length;

  const appUrl = `${window.location.origin}${window.location.pathname}`;
  const defaultSubject = `New survey: ${survey.name}`;
  const defaultMessage = `Hi,\n\nA new survey, "${survey.name}", is now available for you to complete.\n\nSign in at ${appUrl} and go to the Surveys tab to answer it.\n\nThanks!`;

  const [subject, setSubject] = useState(defaultSubject);
  const [message, setMessage] = useState(defaultMessage);
  const [sending, setSending] = useState(false);
  const [result, setResult] = useState(null);

  const handleSend = async () => {
    setSending(true);
    setResult(null);
    try {
      const res = await fetch("/api/send-reminder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ recipients: recipients.map((p) => p.email.trim()), subject, message }),
      });
      const data = await res.json();
      if (!res.ok) {
        setResult({ ok: false, error: data.error || "Failed to send." });
      } else {
        setResult({ ok: true });
        showToast(`Notified ${recipients.length} participant${recipients.length === 1 ? "" : "s"}`);
      }
    } catch (err) {
      setResult({ ok: false, error: "Network error while sending." });
    } finally {
      setSending(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Send to Participants</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>

        <div className="mb-4 flex items-center gap-2 rounded-lg bg-blue-50 px-3 py-2.5 text-sm text-blue-700">
          <Mail size={15} /> Notifying {recipients.length} participant{recipients.length === 1 ? "" : "s"} in "{formatAudience(survey.audience)}" that this survey is ready.
        </div>
        {skippedCount > 0 && (
          <div className="mb-4 flex items-center gap-2 rounded-lg bg-amber-50 px-3 py-2.5 text-xs text-amber-700">
            <XCircle size={13} /> Skipping {skippedCount} participant{skippedCount === 1 ? "" : "s"} with a missing or invalid email address.
          </div>
        )}

        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Subject</label>
            <input value={subject} onChange={(e) => setSubject(e.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400" />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Message</label>
            <textarea rows={6} value={message} onChange={(e) => setMessage(e.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400" />
          </div>
        </div>

        {result && !result.ok && (
          <div className="mt-3 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">
            {result.error}
            {String(result.error || "").includes("RESEND_API_KEY") && (
              <div className="mt-1">Make sure <code>RESEND_API_KEY</code> is set in this project's Vercel environment variables, then redeploy.</div>
            )}
          </div>
        )}
        {result && result.ok && (
          <div className="mt-3 flex items-center gap-1.5 rounded-lg bg-emerald-50 px-3 py-2.5 text-xs text-emerald-700">
            <CheckCircle2 size={13} /> Sent successfully.
          </div>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Close
          </button>
          <button
            disabled={sending || recipients.length === 0}
            onClick={handleSend}
            className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {sending ? <RefreshCw size={14} className="animate-spin" /> : <Mail size={14} />}
            {sending ? "Sending…" : "Send"}
          </button>
        </div>
      </div>
    </div>
  );
}

function ShareSurveyModal({ survey, onClose, onUpdateSurvey }) {
  const [copied, setCopied] = useState(false);
  const url = `${window.location.origin}${window.location.pathname}?survey=${survey.id}`;
  const logoInputRef = React.useRef(null);

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(url);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      // clipboard API can fail without permission — link is still shown to copy manually
    }
  };

  const handleLogoUpload = (file) => {
    if (!file || !onUpdateSurvey) return;
    const reader = new FileReader();
    reader.onload = () => onUpdateSurvey(survey.id, { logoUrl: reader.result });
    reader.readAsDataURL(file);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-sm rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Share "{survey.name}"</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>
        <div className="flex flex-col items-center rounded-xl border border-slate-100 bg-slate-50 p-5">
          <QRCodeImage value={url} size={200} />
          <p className="mt-3 text-center text-xs text-slate-500">Scan to open the survey — no login required.</p>
        </div>
        {onUpdateSurvey && (
          <div className="mt-3 flex items-center justify-between gap-2 rounded-lg bg-slate-50 px-3 py-2">
            <div className="flex items-center gap-2">
              <ClientLogo logoUrl={survey.logoUrl} size={28} radius={6} />
              <span className="text-xs text-slate-500">{survey.logoUrl ? "Client logo set" : "Using Mercer logo (default)"}</span>
            </div>
            <input
              ref={logoInputRef}
              type="file"
              accept="image/*"
              className="hidden"
              onChange={(e) => handleLogoUpload(e.target.files && e.target.files[0])}
            />
            <button
              onClick={() => logoInputRef.current && logoInputRef.current.click()}
              className="flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-blue-700 hover:bg-blue-50"
            >
              <UploadCloud size={12} /> {survey.logoUrl ? "Change" : "Set Logo"}
            </button>
          </div>
        )}
        <div className="mt-4">
          <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Shareable link</label>
          <div className="flex items-center gap-2">
            <input readOnly value={url} className="flex-1 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-600" />
            <button onClick={copyLink} className="flex items-center gap-1 rounded-lg bg-blue-800 px-3 py-2 text-xs font-semibold text-white hover:bg-blue-900">
              {copied ? <CheckCircle2 size={13} /> : <Copy size={13} />} {copied ? "Copied" : "Copy"}
            </button>
          </div>
        </div>
        <div className="mt-4 flex justify-end">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Done
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Survey comparison — side-by-side view of any two surveys' results,
// including live-collected data for admin-built surveys.
// ---------------------------------------------------------------------------

// Detects matching questions between two surveys' analyses (by exact question
// text) and computes the % who Agree/Strongly Agree for each, so two related
// surveys — e.g. a pre-program and post-program run of the same questions —
// can be compared statement-by-statement like a training impact report.
// Counts statements present in one survey but not matched in the other —
// kept separate from matchQuestionsAcrossSurveys so its existing return
// shape (used elsewhere in the UI) doesn't change.
function countUnmatchedStatements(analysisA, analysisB) {
  if (!analysisA?.distributionData?.length || !analysisB?.distributionData?.length) return { onlyInA: 0, onlyInB: 0 };
  const normalize = (s) => (s || "").trim().toLowerCase().replace(/\s+/g, " ");
  const setA = new Set(analysisA.distributionData.map((r) => normalize(r.fullQuestion || r.question)));
  const setB = new Set(analysisB.distributionData.map((r) => normalize(r.fullQuestion || r.question)));
  let onlyInA = 0, onlyInB = 0;
  setA.forEach((q) => { if (!setB.has(q)) onlyInA++; });
  setB.forEach((q) => { if (!setA.has(q)) onlyInB++; });
  return { onlyInA, onlyInB };
}

function matchQuestionsAcrossSurveys(analysisA, analysisB) {
  if (!analysisA?.distributionData?.length || !analysisB?.distributionData?.length) return [];
  const normalize = (s) => (s || "").trim().toLowerCase().replace(/\s+/g, " ");
  const agreePct = (row) => Math.round((row["Agree"] || 0) + (row["Strongly Agree"] || 0));

  const mapB = new Map();
  analysisB.distributionData.forEach((row) => mapB.set(normalize(row.fullQuestion || row.question), row));

  const matches = [];
  analysisA.distributionData.forEach((rowA) => {
    const rowB = mapB.get(normalize(rowA.fullQuestion || rowA.question));
    if (!rowB) return;
    const pre = agreePct(rowA);
    const post = agreePct(rowB);
    matches.push({
      statement: rowA.fullQuestion || rowA.question,
      pre,
      post,
      impact: post - pre,
      band: post > pre ? "increase" : post < pre ? "decrease" : "same",
    });
  });
  return matches.sort((a, b) => b.impact - a.impact);
}

function ComparisonDonut({ data, title }) {
  if (!data || data.length === 0) {
    return <div className="flex h-40 items-center justify-center text-sm text-slate-400">No sentiment data yet.</div>;
  }
  const total = data.reduce((s, d) => s + d.value, 0);
  return (
    <div>
      <div className="flex items-center justify-center gap-6">
        <div className="relative shrink-0" style={{ width: 150 }}>
          <ResponsiveContainer width="100%" height={150}>
            <PieChart>
              <Pie data={data} dataKey="value" nameKey="name" innerRadius={42} outerRadius={65} paddingAngle={2}>
                {data.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Pie>
              <Tooltip formatter={(value, name) => [`${value} (${Math.round((value / total) * 100)}%)`, name]} />
            </PieChart>
          </ResponsiveContainer>
          <div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center">
            <div className="text-lg font-bold text-slate-800">{total}</div>
            <div className="text-xs text-slate-400">responses</div>
          </div>
        </div>
        <div className="space-y-1.5">
          {data.map((d) => (
            <div key={d.name} className="flex items-center gap-2 text-xs">
              <span className="h-2 w-2 shrink-0 rounded-full" style={{ backgroundColor: d.color }} />
              <span className="text-slate-500">{d.name}</span>
              <span className="ml-auto font-semibold text-slate-700">{Math.round((d.value / total) * 100)}%</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function SurveyComparisonView({ surveys, onBack }) {
  const [idA, setIdA] = useState(surveys[0]?.id || "");
  const [idB, setIdB] = useState(surveys[1]?.id || surveys[0]?.id || "");
  const [pptxGenerating, setPptxGenerating] = useState(false);
  const { showToast, projects, currentProjectId } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId);

  const surveyA = surveys.find((s) => s.id === idA);
  const surveyB = surveys.find((s) => s.id === idB);
  const { analysis: liveA } = useLiveSurveyAnalysis(surveyA);
  const { analysis: liveB } = useLiveSurveyAnalysis(surveyB);

  const getAnalysis = (survey, live) =>
    survey?.analysis ?? surveyAnalysisSeed[survey?.id] ?? (survey?.sourceType === "built" ? live : undefined);

  const analysisA = getAnalysis(surveyA, liveA);
  const analysisB = getAnalysis(surveyB, liveB);
  const matches = matchQuestionsAcrossSurveys(analysisA, analysisB);
  const increases = matches.filter((m) => m.band === "increase").length;
  const decreases = matches.filter((m) => m.band === "decrease").length;
  const same = matches.filter((m) => m.band === "same").length;

  const renderColumn = (survey, label, analysis, setId) => {
    if (!survey) {
      return (
        <div className="flex items-center justify-center rounded-2xl border border-dashed border-slate-200 p-10 text-center text-slate-400">
          No survey selected
        </div>
      );
    }
    const ratingData = analysis?.ratingData ?? surveyRatingData;
    const choiceData = analysis?.choiceData ?? surveyChoiceData;
    const avgRating = ratingData.length ? (ratingData.reduce((s, r) => s + r.avg, 0) / ratingData.length).toFixed(1) : "—";
    return (
      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        <div className="flex flex-wrap items-start justify-between gap-4">
          <div style={{ flex: "1 1 260px" }}>
            <div className="mb-1 text-xs font-semibold uppercase tracking-wide text-slate-400">{label}</div>
            <select
              value={survey.id}
              onChange={(e) => setId(e.target.value)}
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold text-slate-800"
            >
              {surveys.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </div>
          <div className="grid grid-cols-3 gap-3 text-center">
            <div className="rounded-lg bg-slate-50 px-4 py-2">
              <div className="text-lg font-bold text-slate-800">{survey.responses}</div>
              <div className="text-xs text-slate-400">Responses</div>
            </div>
            <div className="rounded-lg bg-slate-50 px-4 py-2">
              <div className="text-lg font-bold text-slate-800">{survey.rate}%</div>
              <div className="text-xs text-slate-400">Rate</div>
            </div>
            <div className="rounded-lg bg-slate-50 px-4 py-2">
              <div className="text-lg font-bold text-emerald-600">{avgRating}</div>
              <div className="text-xs text-slate-400">Avg / 5</div>
            </div>
          </div>
        </div>

        <div className="mt-5 flex flex-wrap items-start gap-6 border-t border-slate-100 pt-5">
          <div style={{ flex: "1 1 320px" }}>
            <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Sentiment distribution</div>
            <ComparisonDonut data={choiceData} />
          </div>

          <div style={{ flex: "1 1 380px" }}>
            <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Ratings by question <span className="font-normal normal-case text-slate-400">(hover a bar for the full text)</span></div>
            {ratingData.length > 0 ? (
              <div className="flex gap-3">
                <ResponsiveContainer width="100%" height={Math.max(160, Math.min(ratingData.length, 8) * 32)}>
                  <BarChart data={ratingData.slice(0, 8).map((r, i) => ({ ...r, label: `Q${i + 1}` }))} layout="vertical" margin={{ top: 4, right: 12, bottom: 4, left: 6 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" horizontal={false} />
                    <XAxis type="number" domain={[0, 5]} tickCount={6} tick={{ fontSize: 10, fill: "#64748b" }} />
                    <YAxis type="category" dataKey="label" width={28} tick={{ fontSize: 10, fill: "#334155" }} />
                    <Tooltip content={<SurveyQuestionTooltip mode="avg" />} />
                    <Bar dataKey="avg" fill="#0d9488" radius={[0, 6, 6, 0]} barSize={16} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            ) : (
              <div className="py-6 text-center text-sm text-slate-400">No rating data yet.</div>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="mx-auto max-w-7xl space-y-6 px-6 py-6">
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <LogoWatermark size={200} className="-right-8 -bottom-12" />
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <h1 className="text-2xl font-bold">Compare Surveys</h1>
            <p className="text-sm text-white/70">Compare two surveys' results, statement by statement</p>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={async () => {
                if (!surveyA || !surveyB) return;
                setPptxGenerating(true);
                try {
                  await generateComparisonPPTX(surveyA, analysisA, surveyB, analysisB, currentProject?.logoUrl);
                } catch (err) {
                  console.error("Comparison PPTX generation failed:", err);
                  showToast("Couldn't generate the slide deck. Please try again.");
                } finally {
                  setPptxGenerating(false);
                }
              }}
              disabled={!surveyA || !surveyB || pptxGenerating}
              className="flex items-center gap-1.5 rounded-lg bg-white/15 px-4 py-2.5 text-sm font-semibold text-white hover:bg-white/25 disabled:cursor-not-allowed disabled:opacity-50"
            >
              {pptxGenerating ? <RefreshCw size={15} className="animate-spin" /> : <FileType size={15} />}
              {pptxGenerating ? "Building slides…" : "Download Slides (PPTX)"}
            </button>
            <button onClick={onBack} className="flex items-center gap-1.5 rounded-lg bg-white px-4 py-2.5 text-sm font-semibold hover:bg-white/90" style={{ color: BRAND_NAVY }}>
              <ChevronRight size={15} style={{ transform: "scaleX(-1)" }} /> Back
            </button>
          </div>
        </div>
      </div>

      <div className="flex flex-col gap-6">
        {renderColumn(surveyA, "Survey A · Pre", analysisA, setIdA)}
        {renderColumn(surveyB, "Survey B · Post", analysisB, setIdB)}
      </div>

      {matches.length > 0 && (
        <div className="rounded-2xl border border-slate-200 bg-white p-6">
          <div className="mb-1 flex items-center justify-between">
            <h2 className="text-lg font-bold text-slate-800">Pre &amp; Post Impact</h2>
            <div className="flex items-center gap-4 text-xs">
              <span className="flex items-center gap-1.5 text-slate-500"><span className="h-2.5 w-2.5 rounded-full bg-emerald-500" /> Increase</span>
              <span className="flex items-center gap-1.5 text-slate-500"><span className="h-2.5 w-2.5 rounded-full bg-red-500" /> Decrease</span>
              <span className="flex items-center gap-1.5 text-slate-500"><span className="h-2.5 w-2.5 rounded-full bg-slate-300" /> No change</span>
            </div>
          </div>
          <p className="mb-4 text-sm text-slate-400">
            {matches.length} matching statement{matches.length === 1 ? "" : "s"} found between these two surveys · % shown is respondents who selected Agree or Strongly Agree
          </p>

          <div className="mb-4 flex flex-wrap gap-3 text-sm">
            {increases > 0 && (
              <span className="rounded-full bg-emerald-50 px-3 py-1.5 font-medium text-emerald-700">
                {increases} statement{increases === 1 ? "" : "s"} increased
              </span>
            )}
            {decreases > 0 && (
              <span className="rounded-full bg-red-50 px-3 py-1.5 font-medium text-red-600">
                {decreases} statement{decreases === 1 ? "" : "s"} decreased
              </span>
            )}
            {same > 0 && (
              <span className="rounded-full bg-slate-100 px-3 py-1.5 font-medium text-slate-600">
                {same} statement{same === 1 ? "" : "s"} unchanged
              </span>
            )}
          </div>

          <div style={{ maxHeight: 480, overflowY: "auto" }}>
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-100 text-xs uppercase tracking-wide text-slate-400">
                  <th className="py-2 text-center font-semibold" style={{ width: 70 }}>Pre</th>
                  <th className="py-2 text-left font-semibold">Behavioral Statement</th>
                  <th className="py-2 text-center font-semibold" style={{ width: 70 }}>Post</th>
                  <th className="py-2 text-center font-semibold" style={{ width: 90 }}>Impact</th>
                </tr>
              </thead>
              <tbody>
                {matches.map((m, i) => (
                  <tr key={i} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                    <td className="py-2.5 text-center font-semibold text-slate-700">{m.pre}%</td>
                    <td className="py-2.5 pr-3 text-slate-700">{m.statement}</td>
                    <td className="py-2.5 text-center font-semibold text-slate-700">{m.post}%</td>
                    <td className="py-2.5 text-center">
                      <span
                        className={`rounded-full px-2 py-0.5 text-xs font-semibold ${
                          m.band === "increase" ? "bg-emerald-50 text-emerald-700" : m.band === "decrease" ? "bg-red-50 text-red-600" : "bg-slate-100 text-slate-500"
                        }`}
                      >
                        {m.band === "same" ? "0%" : `${m.impact > 0 ? "+" : ""}${m.impact}%`}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function SurveyAnalytics() {
  const { showToast, currentProjectId, projects } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId);
  const [surveys, setSurveys] = useState(surveysSeed);
  const [surveyLoaded, setSurveyLoaded] = useState(false);
  const [surveyId, setSurveyId] = useState(surveysSeed[0].id);
  const [cohortFilter, setCohortFilter] = useState("All Cohorts");
  const [showAddModal, setShowAddModal] = useState(false);
  const [confirmDeleteSurvey, setConfirmDeleteSurvey] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const result = await storage.get(SURVEYS_STORAGE_KEY);
        if (!cancelled && result && result.value) {
          const parsed = JSON.parse(result.value);
          if (Array.isArray(parsed) && parsed.length > 0) setSurveys(parsed);
        }
      } catch (err) {
        // nothing saved yet
      } finally {
        if (!cancelled) setSurveyLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!surveyLoaded) return;
    (async () => {
      try {
        await storage.set(SURVEYS_STORAGE_KEY, JSON.stringify(surveys));
      } catch (err) {
        console.error("Failed to save surveys", err);
        showToast("Couldn't save — check the database connection (see README)");
      }
    })();
  }, [surveys, surveyLoaded]);

  const projectSurveys = surveys.filter((s) => (s.projectId || "default") === currentProjectId);
  const survey = projectSurveys.find((s) => s.id === surveyId) || projectSurveys[0];
  const { analysis: liveAnalysis } = useLiveSurveyAnalysis(survey);
  const analysis = survey?.analysis ?? surveyAnalysisSeed[survey?.id] ?? (survey?.sourceType === "built" ? liveAnalysis : undefined);
  const ratingData = analysis?.ratingData ?? surveyRatingData;
  const ratingLabel = analysis?.ratingLabel ?? "Rating Questions — Average Score";
  const choiceData = analysis?.choiceData ?? surveyChoiceData;
  const choiceTitle = analysis ? analysis.choiceTitle : "Would you recommend this program?";
  const choiceSubtitle = analysis?.choiceSubtitle ?? null;
  const keywords = analysis ? analysis.keywords : surveyKeywords;
  const avgRating = ratingData.length ? (ratingData.reduce((s, r) => s + r.avg, 0) / ratingData.length).toFixed(1) : "—";
  const [compareMode, setCompareMode] = useState(false);
  const [showShareModal, setShowShareModal] = useState(false);
  const [showSendModal, setShowSendModal] = useState(false);
  const [pptxGenerating, setPptxGenerating] = useState(false);

  if (compareMode) {
    return <SurveyComparisonView surveys={projectSurveys} onBack={() => setCompareMode(false)} />;
  }

  const handleCreateSurvey = (data) => {
    setSurveys((prev) => [{ ...data, projectId: currentProjectId }, ...prev]);
    setSurveyId(data.id);
    setShowAddModal(false);
    showToast(data.sourceType === "send" ? `Sent "${data.name}" to participants` : `Added "${data.name}"`);
  };

  const updateSurvey = (id, data) => {
    setSurveys((prev) => prev.map((s) => (s.id === id ? { ...s, ...data } : s)));
  };

  const deleteSurvey = (id) => {
    const target = surveys.find((s) => s.id === id);
    setSurveys((prev) => {
      const next = prev.filter((s) => s.id !== id);
      if (id === surveyId) setSurveyId(next.length > 0 ? next[0].id : "");
      return next;
    });
    showToast(`Deleted "${target?.name ?? "survey"}"`);
  };

  if (!survey) {
    return (
      <div className="mx-auto max-w-7xl space-y-6 px-6 py-6">
        <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
          <LogoWatermark size={200} className="-right-8 -bottom-12" />
          <div className="flex flex-wrap items-start justify-between gap-6">
            <div className="flex items-start gap-4">
              <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/15"><PieChartIcon size={22} /></div>
              <div>
                <h1 className="text-2xl font-bold">Survey Analytics</h1>
                <p className="text-sm text-white/70">Upload survey results and explore response patterns across cohorts</p>
              </div>
            </div>
            <button onClick={() => setShowAddModal(true)} className="flex items-center gap-1.5 rounded-lg bg-white px-4 py-2.5 text-sm font-semibold hover:bg-white/90" style={{ color: BRAND_NAVY }}>
              <Plus size={15} /> Add Survey
            </button>
          </div>
        </div>
        <div className="flex flex-col items-center justify-center rounded-2xl border border-dashed border-slate-200 bg-white py-16 text-center text-slate-400">
          <PieChartIcon size={28} className="mb-2" />
          <p className="text-sm">No surveys yet. Add one to see analytics here.</p>
        </div>
        {showAddModal && <AddSurveyModal onClose={() => setShowAddModal(false)} onCreate={handleCreateSurvey} />}
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-7xl space-y-6 px-6 py-6">
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <LogoWatermark size={200} className="-right-8 -bottom-12" />
        <div className="flex flex-wrap items-start justify-between gap-6">
          <div className="flex items-start gap-4">
            <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/15"><PieChartIcon size={22} /></div>
            <div>
              <h1 className="text-2xl font-bold">Survey Analytics</h1>
              <p className="text-sm text-white/70">Upload survey results and explore response patterns across cohorts</p>
            </div>
          </div>
          <div className="flex flex-wrap gap-2">
            <button
              onClick={() => setCompareMode(true)}
              disabled={surveys.length < 2}
              className="flex items-center gap-1.5 rounded-lg bg-white/15 px-4 py-2.5 text-sm font-semibold hover:bg-white/25 disabled:cursor-not-allowed disabled:opacity-40"
            >
              <BarChart3 size={15} /> Compare Surveys
            </button>
            <button onClick={() => setShowAddModal(true)} className="flex items-center gap-1.5 rounded-lg bg-white px-4 py-2.5 text-sm font-semibold hover:bg-white/90" style={{ color: BRAND_NAVY }}>
              <Plus size={15} /> Add Survey
            </button>
          </div>
        </div>
      </div>

      {/* Survey picker + filters */}
      <div className="flex flex-wrap items-center gap-3 rounded-2xl border border-slate-200 bg-white p-4">
        <select
          value={surveyId}
          onChange={(e) => setSurveyId(e.target.value)}
          className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-700"
        >
          {projectSurveys.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
        <button
          onClick={() => setConfirmDeleteSurvey(true)}
          title="Delete this survey"
          className="flex items-center gap-1 rounded-lg border border-slate-200 px-2.5 py-2 text-sm text-slate-400 hover:border-red-200 hover:text-red-600"
        >
          <Trash2 size={14} />
        </button>
        {survey.sourceType === "built" && (
          <button
            onClick={() => setShowShareModal(true)}
            className="flex items-center gap-1.5 rounded-lg bg-blue-50 px-3 py-2 text-sm font-medium text-blue-700 hover:bg-blue-100"
          >
            <QrCode size={14} /> Share Survey
          </button>
        )}
        {survey.sourceType === "built" && (
          <button
            onClick={() => setShowSendModal(true)}
            className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
          >
            <Mail size={14} /> Send to Participants
          </button>
        )}
        <button
          onClick={() => downloadSurveyAnalysisCSV(survey, analysis)}
          disabled={!analysis}
          className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40"
        >
          <Download size={14} /> Download Analysis (CSV)
        </button>
        <button
          onClick={async () => {
            setPptxGenerating(true);
            try {
              await generateSurveyPPTX(survey, analysis, currentProject?.logoUrl);
            } catch (err) {
              console.error("PPTX generation failed:", err);
              showToast("Couldn't generate the slide deck. Please try again.");
            } finally {
              setPptxGenerating(false);
            }
          }}
          disabled={!analysis || pptxGenerating}
          className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40"
        >
          {pptxGenerating ? <RefreshCw size={14} className="animate-spin" /> : <FileType size={14} />}
          {pptxGenerating ? "Building slides…" : "Download Slides (PPTX)"}
        </button>
        <span className="text-sm text-slate-400">·</span>
        <div className="flex items-center gap-2 text-sm text-slate-500">
          <Clock size={14} />
          {survey.sourceType === "send" ? "Sent" : survey.sourceType === "built" ? "Created" : "Uploaded"} {survey.uploaded}
          {(survey.sourceType === "send" || survey.sourceType === "built") && survey.responses === 0 && (
            <span className="rounded-full bg-amber-50 px-2 py-0.5 text-xs font-medium text-amber-600">Awaiting responses</span>
          )}
        </div>
        <div className="ml-auto flex flex-wrap items-center gap-2">
          <Filter size={14} className="text-slate-400" />
          <select value={cohortFilter} onChange={(e) => setCohortFilter(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm text-slate-600">
            <option>All Cohorts</option>
            <option>Cohort A</option>
            <option>Cohort B</option>
            <option>Cohort C</option>
          </select>
          <select className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm text-slate-600">
            <option>All Time</option>
            <option>Last 30 days</option>
            <option>Last quarter</option>
          </select>
          <button className="flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-1.5 text-sm text-slate-500 hover:bg-slate-50">
            <RefreshCw size={13} /> Reset
          </button>
        </div>
      </div>

      {showAddModal && <AddSurveyModal onClose={() => setShowAddModal(false)} onCreate={handleCreateSurvey} />}
      {showShareModal && <ShareSurveyModal survey={survey} onClose={() => setShowShareModal(false)} onUpdateSurvey={updateSurvey} />}
      {showSendModal && <SendSurveyModal survey={survey} onClose={() => setShowSendModal(false)} />}
      {confirmDeleteSurvey && (
        <ConfirmModal
          title="Delete this survey?"
          message={`"${survey.name}" and all of its analysis will be permanently removed. This can't be undone.`}
          onCancel={() => setConfirmDeleteSurvey(false)}
          onConfirm={() => {
            deleteSurvey(survey.id);
            setConfirmDeleteSurvey(false);
          }}
        />
      )}

      {/* Summary stats */}
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        <StatCard icon={Users} label="Responses" value={survey.responses} sub={`${cohortFilter}`} iconBg="bg-blue-50" iconColor="text-blue-600" />
        <StatCard icon={Percent} label="Response Rate" value={`${survey.rate}%`} sub="of invited participants" iconBg="bg-teal-50" iconColor="text-teal-600" />
        <StatCard icon={TrendingUp} label="Avg. Rating" value={`${avgRating} / 5`} sub="across rating questions" iconBg="bg-emerald-50" iconColor="text-emerald-600" />
        <StatCard icon={MessageSquare} label="Open Responses" value={keywords ? "118" : "0"} sub={keywords ? "free-text answers" : "no open-text questions"} iconBg="bg-amber-50" iconColor="text-amber-600" />
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        <h2 className="text-lg font-bold text-slate-800">{ratingLabel}</h2>
        <p className="mb-2 text-sm text-slate-400">Likert-scale questions, 1–5 · hover a bar for the full question</p>
        <ResponsiveContainer width="100%" height={Math.max(260, ratingData.length * 34)}>
          <BarChart data={ratingData} layout="vertical" margin={{ top: 8, right: 16, bottom: 8, left: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" horizontal={false} />
            <XAxis type="number" domain={[0, 5]} tickCount={6} tick={{ fontSize: 11, fill: "#64748b" }} />
            <YAxis type="category" dataKey="question" width={260} tick={{ fontSize: 11, fill: "#334155" }} />
            <Tooltip content={<SurveyQuestionTooltip mode="avg" />} />
            <Bar dataKey="avg" fill="#0d9488" radius={[0, 6, 6, 0]} barSize={18} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        <h2 className="text-lg font-bold text-slate-800">{choiceTitle}</h2>
        <p className="mb-4 text-sm text-slate-400">{choiceSubtitle ?? `Multiple choice — ${choiceData.reduce((s, d) => s + d.value, 0)} responses`}</p>
        <div className="flex flex-wrap items-center justify-center gap-8">
          <div className="relative shrink-0" style={{ width: 190 }}>
            <ResponsiveContainer width="100%" height={190}>
              <PieChart>
                <Pie data={choiceData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={82} paddingAngle={2}>
                  {choiceData.map((d, i) => <Cell key={i} fill={d.color} />)}
                </Pie>
                <Tooltip formatter={(value, name) => [`${value} (${Math.round((value / choiceData.reduce((s, d) => s + d.value, 0)) * 100)}%)`, name]} />
              </PieChart>
            </ResponsiveContainer>
            <div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center">
              <div className="text-2xl font-bold text-slate-800">{choiceData.reduce((s, d) => s + d.value, 0)}</div>
              <div className="text-xs text-slate-400">responses</div>
            </div>
          </div>
          <div className="w-full" style={{ maxWidth: 260 }}>
            {choiceData.map((d) => {
              const total = choiceData.reduce((s, x) => s + x.value, 0);
              const pct = Math.round((d.value / total) * 100);
              return (
                <div key={d.name} className="mb-3 last:mb-0">
                  <div className="mb-1 flex items-center justify-between text-sm">
                    <span className="flex items-center gap-2 text-slate-600">
                      <span className="h-2.5 w-2.5 shrink-0 rounded-full" style={{ backgroundColor: d.color }} />
                      {d.name}
                    </span>
                    <span className="font-semibold text-slate-800">{pct}%</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <div className="h-1.5 flex-1 rounded-full bg-slate-100">
                      <div className="h-1.5 rounded-full" style={{ width: `${pct}%`, backgroundColor: d.color }} />
                    </div>
                    <span className="text-xs text-slate-400" style={{ width: 26, textAlign: "right" }}>{d.value}</span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {analysis?.distributionData && (
        <div className="rounded-2xl border border-slate-200 bg-white p-6">
          <h2 className="text-lg font-bold text-slate-800">Response Distribution by Question</h2>
          <p className="mb-3 text-sm text-slate-400">Every question analyzed independently — % of respondents per answer choice</p>
          <div style={{ maxHeight: 520, overflowY: "auto" }}>
            <ResponsiveContainer width="100%" height={Math.max(220, analysis.distributionData.length * 30)}>
              <BarChart data={analysis.distributionData} layout="vertical" margin={{ top: 8, right: 16, bottom: 8, left: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" horizontal={false} />
                <XAxis type="number" domain={[0, 100]} tickCount={6} tickFormatter={(v) => `${v}%`} tick={{ fontSize: 11, fill: "#64748b" }} />
                <YAxis type="category" dataKey="question" width={260} tick={{ fontSize: 11, fill: "#334155" }} />
                <Tooltip content={<SurveyQuestionTooltip mode="dist" />} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                {LIKERT_LABELS_5.map((label, i) => (
                  <Bar key={label} dataKey={label} stackId="dist" fill={LIKERT_COLORS_5[i]} />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {analysis?.fullQuestionList && (
        <div className="rounded-2xl border border-slate-200 bg-white p-6">
          <div className="mb-3 flex items-center justify-between">
            <div>
              <h2 className="text-lg font-bold text-slate-800">All Questions — Full Text</h2>
              <p className="text-sm text-slate-400">Exactly as written in the source file, ranked by average score</p>
            </div>
            <span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs font-medium text-slate-500">{analysis.fullQuestionList.length} questions</span>
          </div>
          <div style={{ maxHeight: 420, overflowY: "auto" }}>
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
                  <th className="py-2 pr-3 font-semibold" style={{ width: 44 }}>#</th>
                  <th className="py-2 pr-3 font-semibold">Statement</th>
                  <th className="py-2 text-center font-semibold" style={{ width: 90 }}>Avg</th>
                  <th className="py-2 text-center font-semibold" style={{ width: 90 }}>Responses</th>
                </tr>
              </thead>
              <tbody>
                {analysis.fullQuestionList.map((q, i) => (
                  <tr key={i} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                    <td className="py-2.5 pr-3 text-slate-400">{i + 1}</td>
                    <td className="py-2.5 pr-3 text-slate-700">{q.statement}</td>
                    <td className={`py-2.5 text-center font-semibold ${q.avg >= 4 ? "text-emerald-600" : q.avg >= 3 ? "text-amber-600" : "text-red-600"}`}>{q.avg.toFixed(2)}</td>
                    <td className="py-2.5 text-center text-slate-400">{q.n}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="flex flex-wrap items-start gap-6">
        <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "2 1 420px" }}>
          <div className="mb-3 flex items-center justify-between">
            <h2 className="text-lg font-bold text-slate-800">{analysis ? analysis.crossTabTitle : "Cross-Tabulation — Rating by Cohort"}</h2>
            <div className="flex gap-2">
              <button className="flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-500 hover:bg-slate-50"><Download size={12} /> CSV</button>
              <button className="flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-500 hover:bg-slate-50"><Download size={12} /> PNG</button>
            </div>
          </div>
          {analysis ? (
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-100 text-xs uppercase tracking-wide text-slate-400">
                  <th className="py-2 text-left font-semibold">Statement</th>
                  <th className="py-2 text-center font-semibold" style={{ width: 110 }}>Avg Score</th>
                </tr>
              </thead>
              <tbody>
                {analysis.crossTabRows.map((r, i) => (
                  <tr key={i} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                    <td className="py-2.5 text-left text-slate-700">
                      <span className={`mr-2 rounded-full px-2 py-0.5 text-xs font-medium ${r.band === "top" ? "bg-emerald-50 text-emerald-600" : "bg-amber-50 text-amber-600"}`}>
                        {r.band === "top" ? "Top" : "Growth"}
                      </span>
                      {r.statement}
                    </td>
                    <td className="py-2.5 text-center font-semibold text-slate-800">{r.avg.toFixed(2)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <table className="w-full text-sm" style={{ tableLayout: "fixed" }}>
              <thead>
                <tr className="border-b border-slate-100 text-xs uppercase tracking-wide text-slate-400">
                  <th className="py-2 text-left font-semibold" style={{ width: "22%" }}>Cohort</th>
                  <th className="py-2 text-center font-semibold">Content Quality</th>
                  <th className="py-2 text-center font-semibold">Facilitator</th>
                  <th className="py-2 text-center font-semibold">Pace</th>
                  <th className="py-2 text-center font-semibold">Relevance</th>
                  <th className="py-2 text-center font-semibold">Respondents</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { cohort: "Cohort A", content: 4.5, fac: 4.7, pace: 3.8, rel: 4.3, n: 62 },
                  { cohort: "Cohort B", content: 4.3, fac: 4.5, pace: 4.0, rel: 4.1, n: 74 },
                  { cohort: "Cohort C", content: 4.4, fac: 4.6, pace: 3.9, rel: 4.2, n: 50 },
                ].map((r) => (
                  <tr key={r.cohort} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                    <td className="py-2.5 text-left"><span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-medium text-blue-700">{r.cohort}</span></td>
                    <td className="py-2.5 text-center font-medium text-slate-700">{r.content.toFixed(1)}</td>
                    <td className="py-2.5 text-center font-medium text-slate-700">{r.fac.toFixed(1)}</td>
                    <td className="py-2.5 text-center font-medium text-slate-700">{r.pace.toFixed(1)}</td>
                    <td className="py-2.5 text-center font-medium text-slate-700">{r.rel.toFixed(1)}</td>
                    <td className="py-2.5 text-center text-slate-500">{r.n}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        <div className="rounded-2xl border border-slate-200 bg-white p-6" style={{ flex: "1 1 280px" }}>
          <div className="mb-3 flex items-center gap-2">
            <Tag size={15} className="text-blue-600" />
            <h2 className="text-lg font-bold text-slate-800">Keyword Themes</h2>
          </div>
          {keywords ? (
            <>
              <p className="mb-3 text-sm text-slate-400">Extracted from open-text responses</p>
              <div className="flex flex-wrap gap-2">
                {keywords.map((k) => (
                  <span
                    key={k.word}
                    className="rounded-full bg-slate-100 px-3 py-1.5 text-slate-600"
                    style={{ fontSize: `${11 + k.weight}px`, fontWeight: k.weight >= 4 ? 600 : 500 }}
                  >
                    {k.word}
                  </span>
                ))}
              </div>
              <div className="mt-4 rounded-lg bg-blue-50 px-3 py-2.5 text-xs text-blue-700">
                <strong>Top strengths:</strong> engaging facilitators, actionable takeaways.<br />
                <strong>Areas to improve:</strong> pacing, more case studies.
              </div>
            </>
          ) : (
            <div className="flex flex-col items-center justify-center py-8 text-center text-slate-400">
              <MessageSquare size={22} className="mb-2" />
              <p className="text-sm">This survey has no open-text questions, so there are no keyword themes to extract.</p>
            </div>
          )}
        </div>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-6">
        <div className="mb-3 flex items-center justify-between">
          <div>
            <h2 className="text-lg font-bold text-slate-800">Export & Templates</h2>
            <p className="text-sm text-slate-400">Save this view or export it for a stakeholder deck</p>
          </div>
        </div>
        <div className="flex flex-wrap gap-3">
          <button className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"><Download size={14} /> Export raw data (CSV)</button>
          <button className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"><Download size={14} /> Export charts (PNG)</button>
          <button className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"><FileType size={14} /> Export deck (PPTX-ready)</button>
          <button className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900"><Save size={14} /> Save as report template</button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Import Grading Criteria — reuses the same workbook parser as participant
// import, but instead of importing people, treats every detected column as
// a candidate grading criterion with an editable weight.
// ---------------------------------------------------------------------------

function ImportCriteriaModal({ onClose, onImport }) {
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState(null);
  const [candidates, setCandidates] = useState(null); // [{label, weight, include}]
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = React.useRef(null);

  const handleFile = async (file) => {
    if (!file) return;
    setError(null);
    setParsing(true);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const result = parseParticipantWorkbook(workbook);
      if (!result.sheetName || result.criteriaColumns.length === 0) {
        setError("Couldn't find any candidate grading-criteria columns in this file.");
        setCandidates(null);
      } else {
        const evenWeight = Math.round((100 / result.criteriaColumns.length) * 10) / 10;
        setCandidates(result.criteriaColumns.map((c) => ({ label: c.label, weight: evenWeight, include: true })));
      }
    } catch (err) {
      setError("Couldn't read this file. Make sure it's a valid .xlsx file.");
    } finally {
      setParsing(false);
    }
  };

  const included = (candidates || []).filter((c) => c.include);
  const totalWeight = included.reduce((s, c) => s + Number(c.weight || 0), 0);

  const handleConfirm = () => {
    const newCriteria = included.map((c, i) => ({ id: `${Date.now()}-${i}`, name: c.label, weight: Number(c.weight) || 0, isReflectionCriterion: false }));
    onImport(newCriteria);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-lg rounded-2xl bg-white p-6 shadow-xl" style={{ maxHeight: "85vh", overflowY: "auto" }}>
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Import Grading Criteria from Excel</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>
        <p className="mb-3 text-xs text-slate-400">This replaces the current criteria list for this project. Upload the same kind of tracker file used for participant import — every column becomes a candidate criterion, with an even starting weight you can adjust.</p>

        {!candidates && (
          <>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleFile(e.target.files && e.target.files[0])} />
            <div
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => { e.preventDefault(); setDragOver(false); const file = e.dataTransfer.files && e.dataTransfer.files[0]; if (file) handleFile(file); }}
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-4 py-10 text-center transition-colors ${dragOver ? "border-blue-400 bg-blue-50" : "border-slate-200 bg-slate-50"}`}
            >
              {parsing ? (
                <><RefreshCw size={22} className="mb-2 animate-spin text-blue-500" /><p className="text-sm font-medium text-slate-600">Reading file…</p></>
              ) : (
                <><UploadCloud size={24} className={dragOver ? "text-blue-500" : "text-slate-400"} /><p className="mt-2 text-sm font-medium text-slate-600">Drag &amp; drop an .xlsx file, or click to browse</p></>
              )}
            </div>
            {error && <div className="mt-3 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{error}</div>}
          </>
        )}

        {candidates && (
          <>
            <div className="mb-2 flex items-center justify-between text-xs">
              <span className="text-slate-400">{included.length} of {candidates.length} columns selected</span>
              <span className={`font-semibold ${totalWeight === 100 ? "text-emerald-600" : "text-amber-600"}`}>Total weight: {totalWeight}%</span>
            </div>
            <div className="space-y-1.5" style={{ maxHeight: 320, overflowY: "auto" }}>
              {candidates.map((c, i) => (
                <div key={c.label} className="flex items-center gap-2 rounded-lg bg-slate-50 px-3 py-2">
                  <input type="checkbox" checked={c.include} onChange={(e) => setCandidates((prev) => prev.map((p, pi) => (pi === i ? { ...p, include: e.target.checked } : p)))} />
                  <span className="flex-1 truncate text-xs text-slate-600" title={c.label}>{c.label}</span>
                  <input
                    type="number"
                    min="0"
                    max="100"
                    disabled={!c.include}
                    value={c.weight}
                    onChange={(e) => setCandidates((prev) => prev.map((p, pi) => (pi === i ? { ...p, weight: Number(e.target.value) } : p)))}
                    className="w-16 rounded border border-slate-200 bg-white px-2 py-1 text-xs disabled:opacity-40"
                  />
                  <span className="text-xs text-slate-400">%</span>
                </div>
              ))}
            </div>
            <button onClick={() => setCandidates(null)} className="mt-3 text-xs font-medium text-blue-700 hover:underline">Choose a different file</button>
          </>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">Cancel</button>
          {candidates && (
            <button
              disabled={included.length === 0}
              onClick={handleConfirm}
              className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
            >
              <CheckCircle2 size={14} /> Use These Criteria
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Groups — organize a project's participants into working groups (e.g. for
// group coaching sessions or team assignments), independent of cohort.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Overview — a single at-a-glance view of a project's key numbers, so admin
// doesn't need to click through every tab to get a sense of how things
// stand. Loads meetings and surveys itself (participants/projects already
// come from context) since those aren't otherwise available at this level.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Admin Accounts — manage who has admin access. Only a signed-in admin can
// reach this at all, and only a signed-in admin can create another admin
// account (there's no public "sign up as admin" path).
// ---------------------------------------------------------------------------

function AdminAccountsPanel() {
  const { adminAccounts, currentAdmin, addAdminAccount, removeAdminAccount } = useAppData();
  const [showAdd, setShowAdd] = useState(false);
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState(null);
  const [submitting, setSubmitting] = useState(false);
  const [confirmRemove, setConfirmRemove] = useState(null);

  const resetForm = () => { setName(""); setEmail(""); setPassword(""); setError(null); setShowAdd(false); };

  const handleAdd = async () => {
    setSubmitting(true);
    setError(null);
    try {
      const result = await addAdminAccount(name, email, password);
      if (!result.ok) { setError(result.error); return; }
      resetForm();
    } catch (err) {
      setError("Something went wrong. Please try again.");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-bold text-slate-800">Admin Accounts</h2>
          <p className="text-sm text-slate-400">Who has admin access to this deployment</p>
        </div>
        {!showAdd && (
          <button onClick={() => setShowAdd(true)} className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
            <Plus size={15} /> Add Admin
          </button>
        )}
      </div>

      {showAdd && (
        <div className="mb-4 space-y-3 rounded-xl border border-slate-100 bg-slate-50 p-4">
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Name</label>
              <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Full name" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm" />
            </div>
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Email</label>
              <input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@organization.com" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm" />
            </div>
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Password</label>
            <div className="relative">
              <input
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-2 pr-10 text-sm"
              />
              <button type="button" onClick={() => setShowPassword((s) => !s)} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">
                {showPassword ? <EyeOff size={15} /> : <Eye size={15} />}
              </button>
            </div>
          </div>
          {error && <div className="rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{error}</div>}
          <div className="flex justify-end gap-2">
            <button onClick={resetForm} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">Cancel</button>
            <button
              onClick={handleAdd}
              disabled={submitting || !name.trim() || !email.trim() || !password}
              className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
            >
              {submitting ? <RefreshCw size={14} className="animate-spin" /> : <CheckCircle2 size={14} />}
              Add
            </button>
          </div>
        </div>
      )}

      <div className="space-y-2">
        {adminAccounts.map((a) => (
          <div key={a.id} className="flex items-center justify-between gap-3 rounded-xl border border-slate-100 bg-white px-4 py-3">
            <div className="flex items-center gap-3">
              <div className="flex h-9 w-9 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">{initialsFromName(a.name)}</div>
              <div>
                <div className="flex items-center gap-1.5 text-sm font-medium text-slate-800">
                  {a.name}
                  {currentAdmin && a.id === currentAdmin.id && <span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-semibold text-blue-700">You</span>}
                </div>
                <div className="text-xs text-slate-400">{a.email}</div>
              </div>
            </div>
            <button
              onClick={() => setConfirmRemove(a)}
              disabled={adminAccounts.length <= 1}
              title={adminAccounts.length <= 1 ? "Can't remove the last admin account" : "Remove this admin account"}
              className="text-slate-400 hover:text-red-600 disabled:cursor-not-allowed disabled:opacity-30"
            >
              <Trash2 size={15} />
            </button>
          </div>
        ))}
      </div>

      {confirmRemove && (
        <ConfirmModal
          title={`Remove admin access for ${confirmRemove.name}?`}
          message="They'll no longer be able to sign in as an admin. This can't be undone."
          confirmLabel="Remove"
          onCancel={() => setConfirmRemove(null)}
          onConfirm={() => { removeAdminAccount(confirmRemove.id); setConfirmRemove(null); }}
        />
      )}
    </div>
  );
}

function OverviewPanel() {
  const { participants, projects, currentProjectId } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId) || projects[0];
  const projectParticipants = participants.filter((p) => (p.projectId || "default") === currentProjectId);
  const [meetings, setMeetings] = useState([]);
  const [surveys, setSurveys] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      try {
        const [meetingsResult, surveysResult] = await Promise.all([
          storage.get(ATTENDANCE_STORAGE_KEY).catch(() => null),
          storage.get(SURVEYS_STORAGE_KEY).catch(() => null),
        ]);
        if (cancelled) return;
        const allMeetings = meetingsResult?.value ? JSON.parse(meetingsResult.value) : [];
        const allSurveys = surveysResult?.value ? JSON.parse(surveysResult.value) : [];
        setMeetings(Array.isArray(allMeetings) ? allMeetings.filter((m) => (m.projectId || "default") === currentProjectId) : []);
        setSurveys(Array.isArray(allSurveys) ? allSurveys.filter((s) => (s.projectId || "default") === currentProjectId) : []);
      } catch (err) {
        // nothing saved yet
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [currentProjectId]);

  const today = new Date().toISOString().slice(0, 10);
  const upcomingMeetings = meetings.filter((m) => (m.date || "") >= today);
  const builtSurveys = surveys.filter((s) => s.sourceType === "built");

  const scores = projectParticipants
    .map((p) => computeWeightedTotal(p, currentProject) ?? p.avgScore)
    .filter((s) => s !== null && s !== undefined);
  const avgScore = scores.length ? Math.round(scores.reduce((a, b) => a + b, 0) / scores.length) : null;
  const passCount = scores.filter((s) => s >= passingThreshold).length;
  const passRate = scores.length ? Math.round((passCount / scores.length) * 100) : null;

  const resources = currentProject?.resources || [];
  const totalPossibleCompletions = resources.length * projectParticipants.length;
  const actualCompletions = projectParticipants.reduce((sum, p) => sum + (p.completedResourceIds || []).filter((id) => resources.some((r) => r.id === id)).length, 0);
  const resourceCompletionRate = totalPossibleCompletions > 0 ? Math.round((actualCompletions / totalPossibleCompletions) * 100) : null;

  const groups = currentProject?.groups || [];
  const groupedCount = projectParticipants.filter((p) => groups.some((g) => g.participantIds.includes(p.employeeId))).length;
  const unassignedCount = projectParticipants.length - groupedCount;

  const statCard = (label, value, icon, color) => {
    const Icon = icon;
    return (
      <div className="rounded-2xl border border-slate-200 bg-white p-5">
        <div className="mb-2 flex items-center gap-2">
          <div className="flex h-8 w-8 items-center justify-center rounded-lg" style={{ backgroundColor: `${color}15`, color }}><Icon size={16} /></div>
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">{label}</span>
        </div>
        <div className="text-2xl font-bold text-slate-800">{value}</div>
      </div>
    );
  };

  return (
    <div>
      <div className="mb-5">
        <h2 className="text-lg font-bold text-slate-800">Overview — {currentProject?.name}</h2>
        <p className="text-sm text-slate-400">Key numbers for this project at a glance</p>
      </div>

      {loading ? (
        <div className="flex items-center justify-center py-14 text-slate-400"><RefreshCw size={22} className="animate-spin" /></div>
      ) : (
        <>
          <div className="mb-6 grid grid-cols-2 gap-4 sm:grid-cols-3 lg:grid-cols-4">
            {statCard("Participants", projectParticipants.length, Users, "#1d4ed8")
            }
            {statCard("Average Score", avgScore !== null ? `${avgScore}%` : "—", TrendingUp, "#0f9a8e")}
            {statCard("Pass Rate", passRate !== null ? `${passRate}%` : "—", Award, "#7c3aed")}
            {statCard("Upcoming Meetings", upcomingMeetings.length, Calendar, "#f97316")}
            {statCard("Surveys", builtSurveys.length, PieChartIcon, "#0ea5e9")}
            {statCard("Resource Completion", resourceCompletionRate !== null ? `${resourceCompletionRate}%` : "—", FileText, "#dc2626")}
            {statCard("Groups", groups.length, Users, "#65a30d")}
            {statCard("Unassigned to Group", unassignedCount, UserCheck, "#64748b")}
          </div>

          {projectParticipants.length === 0 && (
            <div className="rounded-xl border border-dashed border-slate-200 bg-white px-4 py-6 text-center text-sm text-slate-400">
              This project has no participants yet — add some via Participants → Add Participant or Import from Excel to see real numbers here.
            </div>
          )}
        </>
      )}
    </div>
  );
}

function GroupsPanel() {
  const { projects, currentProjectId, updateProject, participants } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId) || projects[0];
  const groups = currentProject?.groups || [];
  const projectParticipants = participants.filter((p) => (p.projectId || "default") === currentProjectId);
  const [addMode, setAddMode] = useState(null); // null | 'single' | 'bulk'
  const [newGroupName, setNewGroupName] = useState("");
  const [bulkCount, setBulkCount] = useState(4);
  const [expandedGroupId, setExpandedGroupId] = useState(null);
  const [groupSearch, setGroupSearch] = useState("");

  const setGroups = (next) => updateProject(currentProject.id, { groups: next });

  const closeAdd = () => { setAddMode(null); setNewGroupName(""); };

  const addGroup = () => {
    const trimmed = newGroupName.trim();
    if (!trimmed) return;
    const id = `group-${Date.now()}`;
    setGroups([...groups, { id, name: trimmed, participantIds: [] }]);
    closeAdd();
    setExpandedGroupId(id);
  };

  const addBulkGroups = () => {
    const count = Math.max(1, Math.min(50, Number(bulkCount) || 0));
    if (!count) return;
    const existingNumbers = groups
      .map((g) => { const m = g.name.match(/^Group (\d+)$/); return m ? Number(m[1]) : 0; })
      .filter(Boolean);
    let nextNumber = (existingNumbers.length ? Math.max(...existingNumbers) : 0) + 1;
    const newGroups = [];
    for (let i = 0; i < count; i++) {
      newGroups.push({ id: `group-${Date.now()}-${i}`, name: `Group ${nextNumber}`, participantIds: [] });
      nextNumber++;
    }
    setGroups([...groups, ...newGroups]);
    setAddMode(null);
  };

  const renameGroup = (id, name) => setGroups(groups.map((g) => (g.id === id ? { ...g, name } : g)));
  const deleteGroup = (id) => setGroups(groups.filter((g) => g.id !== id));
  const groupOf = (employeeId) => groups.find((g) => g.participantIds.includes(employeeId));

  const toggleMember = (groupId, employeeId) => {
    setGroups(groups.map((g) => {
      if (g.id === groupId) {
        const has = g.participantIds.includes(employeeId);
        return { ...g, participantIds: has ? g.participantIds.filter((id) => id !== employeeId) : [...g.participantIds, employeeId] };
      }
      // A participant belongs to only one group at a time — adding them
      // here removes them from wherever else they were.
      return { ...g, participantIds: g.participantIds.filter((id) => id !== employeeId) };
    }));
  };

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-bold text-slate-800">Groups — {currentProject?.name}</h2>
          <p className="text-sm text-slate-400">Organize participants into working groups, e.g. for group coaching sessions</p>
        </div>
        {addMode === null && (
          <div className="flex items-center gap-2">
            <button onClick={() => setAddMode("single")} className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
              <Plus size={14} /> New Group
            </button>
            <button onClick={() => setAddMode("bulk")} className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
              <Users size={15} /> Create Multiple Groups
            </button>
          </div>
        )}
        {addMode === "single" && (
          <div className="flex items-center gap-1.5">
            <input
              autoFocus
              value={newGroupName}
              onChange={(e) => setNewGroupName(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") addGroup(); }}
              placeholder="Group name…"
              className="rounded-lg border border-slate-200 px-3 py-2 text-sm"
              style={{ width: 180 }}
            />
            <button onClick={addGroup} disabled={!newGroupName.trim()} className="rounded-lg bg-blue-800 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40">
              Create
            </button>
            <button onClick={closeAdd} className="text-slate-400 hover:text-slate-600"><X size={16} /></button>
          </div>
        )}
        {addMode === "bulk" && (
          <div className="flex items-center gap-1.5">
            <span className="text-sm text-slate-500">Create</span>
            <input
              autoFocus
              type="number"
              min="1"
              max="50"
              value={bulkCount}
              onChange={(e) => setBulkCount(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") addBulkGroups(); }}
              className="rounded-lg border border-slate-200 px-3 py-2 text-sm text-center"
              style={{ width: 70 }}
            />
            <span className="text-sm text-slate-500">groups</span>
            <button onClick={addBulkGroups} className="rounded-lg bg-blue-800 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-900">
              Create
            </button>
            <button onClick={() => setAddMode(null)} className="text-slate-400 hover:text-slate-600"><X size={16} /></button>
          </div>
        )}
      </div>
      {addMode === "bulk" && (
        <p className="-mt-2 mb-4 text-xs text-slate-400">Creates "Group 1", "Group 2", etc. — click any of them afterward to rename it or pick its participants.</p>
      )}

      {groups.length > 0 && (
        <div className="mb-4 flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2" style={{ maxWidth: 320 }}>
          <Search size={15} className="text-slate-400" />
          <input value={groupSearch} onChange={(e) => setGroupSearch(e.target.value)} placeholder="Search groups…" className="w-full text-sm outline-none placeholder:text-slate-400" />
        </div>
      )}

      {(() => {
        const visibleGroups = groups.filter((g) => g.name.toLowerCase().includes(groupSearch.toLowerCase()));
        if (groups.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 py-14 text-center text-slate-400">
              <Users size={26} className="mb-2" />
              <p className="text-sm">No groups yet for this project.</p>
            </div>
          );
        }
        if (visibleGroups.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 py-14 text-center text-slate-400">
              <Search size={26} className="mb-2" />
              <p className="text-sm">No groups match "{groupSearch}".</p>
            </div>
          );
        }
        return (
        <div className="space-y-3">
          {visibleGroups.map((g) => {
            const expanded = expandedGroupId === g.id;
            const members = projectParticipants.filter((p) => g.participantIds.includes(p.employeeId));
            return (
              <div key={g.id} className="rounded-xl border border-slate-100 bg-white">
                <div className="flex items-center justify-between gap-2 px-4 py-3">
                  <button onClick={() => setExpandedGroupId(expanded ? null : g.id)} className="flex flex-1 items-center gap-2 text-left">
                    <ChevronDown size={15} className={`text-slate-400 transition-transform ${expanded ? "" : "-rotate-90"}`} />
                    <input
                      value={g.name}
                      onClick={(e) => e.stopPropagation()}
                      onChange={(e) => renameGroup(g.id, e.target.value)}
                      className="rounded border border-transparent bg-transparent px-1 py-0.5 text-sm font-semibold text-slate-800 hover:border-slate-200 focus:border-blue-400 focus:outline-none"
                    />
                    <span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-medium text-blue-700">{members.length} member{members.length === 1 ? "" : "s"}</span>
                  </button>
                  <button onClick={() => deleteGroup(g.id)} className="text-slate-400 hover:text-red-600"><Trash2 size={15} /></button>
                </div>
                {expanded && (
                  <div className="border-t border-slate-100 px-4 py-3">
                    {(() => {
                      // Only shows participants who are either already in this
                      // group or not in any group yet — someone in a different
                      // group needs to be removed there first before they can
                      // be picked here, so they don't clutter every group's list.
                      const availableParticipants = projectParticipants.filter((p) => {
                        const inThisGroup = g.participantIds.includes(p.employeeId);
                        return inThisGroup || !groupOf(p.employeeId);
                      });
                      if (projectParticipants.length === 0) {
                        return <p className="text-xs text-slate-400">No participants in this project yet.</p>;
                      }
                      if (availableParticipants.length === 0) {
                        return <p className="text-xs text-slate-400">Everyone in this project is already assigned to a group.</p>;
                      }
                      return (
                        <div className="grid grid-cols-2 gap-1.5" style={{ maxHeight: 220, overflowY: "auto" }}>
                          {availableParticipants.map((p) => (
                            <label key={p.employeeId} className="flex items-center gap-2 rounded-lg bg-slate-50 px-2.5 py-1.5 text-xs">
                              <input type="checkbox" checked={g.participantIds.includes(p.employeeId)} onChange={() => toggleMember(g.id, p.employeeId)} />
                              <span className="truncate text-slate-600">{p.name}</span>
                            </label>
                          ))}
                        </div>
                      );
                    })()}
                  </div>
                )}
              </div>
            );
          })}
        </div>
        );
      })()}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Resources — documents/videos participants review and mark as completed.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Resource completion reminders — nudges only participants who haven't
// marked a specific resource as completed yet, reusing the same
// Resend-based infrastructure as meeting and survey reminders.
// ---------------------------------------------------------------------------

function SendResourceReminderModal({ resource, projectParticipants, onClose }) {
  const { showToast } = useAppData();
  const incomplete = projectParticipants.filter((p) => !(p.completedResourceIds || []).includes(resource.id));

  const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const recipients = incomplete.filter((p) => p.email && EMAIL_RE.test(p.email.trim()));
  const skippedCount = incomplete.length - recipients.length;

  const appUrl = `${window.location.origin}${window.location.pathname}`;
  const defaultSubject = `Reminder: "${resource.title}" is still pending`;
  const defaultMessage = `Hi,\n\nJust a reminder that "${resource.title}" is still marked incomplete on your Resources tab.\n\nSign in at ${appUrl} and go to Resources to review it and mark it complete.\n\nThanks!`;

  const [subject, setSubject] = useState(defaultSubject);
  const [message, setMessage] = useState(defaultMessage);
  const [sending, setSending] = useState(false);
  const [result, setResult] = useState(null);

  const handleSend = async () => {
    setSending(true);
    setResult(null);
    try {
      const res = await fetch("/api/send-reminder", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ recipients: recipients.map((p) => p.email.trim()), subject, message }),
      });
      const data = await res.json();
      if (!res.ok) {
        setResult({ ok: false, error: data.error || "Failed to send." });
      } else {
        setResult({ ok: true });
        showToast(`Reminded ${recipients.length} participant${recipients.length === 1 ? "" : "s"}`);
      }
    } catch (err) {
      setResult({ ok: false, error: "Network error while sending." });
    } finally {
      setSending(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Send Reminder</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>

        <div className="mb-4 flex items-center gap-2 rounded-lg bg-blue-50 px-3 py-2.5 text-sm text-blue-700">
          <Mail size={15} /> Reminding {recipients.length} participant{recipients.length === 1 ? "" : "s"} who haven't completed "{resource.title}" yet.
        </div>
        {incomplete.length === 0 && (
          <div className="mb-4 flex items-center gap-2 rounded-lg bg-emerald-50 px-3 py-2.5 text-xs text-emerald-700">
            <CheckCircle2 size={13} /> Everyone has already completed this resource — nothing to send.
          </div>
        )}
        {skippedCount > 0 && (
          <div className="mb-4 flex items-center gap-2 rounded-lg bg-amber-50 px-3 py-2.5 text-xs text-amber-700">
            <XCircle size={13} /> Skipping {skippedCount} participant{skippedCount === 1 ? "" : "s"} with a missing or invalid email address.
          </div>
        )}

        <div className="space-y-3">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Subject</label>
            <input value={subject} onChange={(e) => setSubject(e.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400" />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Message</label>
            <textarea rows={6} value={message} onChange={(e) => setMessage(e.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400" />
          </div>
        </div>

        {result && !result.ok && (
          <div className="mt-3 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">
            {result.error}
            {String(result.error || "").includes("RESEND_API_KEY") && (
              <div className="mt-1">Make sure <code>RESEND_API_KEY</code> is set in this project's Vercel environment variables, then redeploy.</div>
            )}
          </div>
        )}
        {result && result.ok && (
          <div className="mt-3 flex items-center gap-1.5 rounded-lg bg-emerald-50 px-3 py-2.5 text-xs text-emerald-700">
            <CheckCircle2 size={13} /> Sent successfully.
          </div>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Close
          </button>
          <button
            disabled={sending || recipients.length === 0}
            onClick={handleSend}
            className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {sending ? <RefreshCw size={14} className="animate-spin" /> : <Mail size={14} />}
            {sending ? "Sending…" : "Send"}
          </button>
        </div>
      </div>
    </div>
  );
}

function ResourcesPanel() {
  const { projects, currentProjectId, updateProject, participants, showToast } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId) || projects[0];
  const resources = currentProject?.resources || [];
  const projectParticipants = participants.filter((p) => (p.projectId || "default") === currentProjectId);
  const [showAdd, setShowAdd] = useState(false);
  const [title, setTitle] = useState("");
  const [type, setType] = useState("document");
  const [mode, setMode] = useState("upload"); // 'upload' | 'link'
  const [url, setUrl] = useState("");
  const [fileName, setFileName] = useState("");
  const [uploading, setUploading] = useState(false);
  const [uploadError, setUploadError] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [reminderResource, setReminderResource] = useState(null);
  const [resourceSearch, setResourceSearch] = useState("");
  const fileInputRef = React.useRef(null);

  const setResources = (next) => updateProject(currentProject.id, { resources: next });

  const resetForm = () => {
    setTitle(""); setUrl(""); setFileName(""); setType("document"); setMode("upload"); setShowAdd(false); setUploadError(null);
  };

  const handleFileUpload = async (file) => {
    if (!file) return;
    setUploadError(null);
    setUploading(true);
    setFileName(file.name);
    if (!title.trim()) setTitle(file.name.replace(/\.[^.]+$/, ""));
    try {
      const arrayBuffer = await file.arrayBuffer();
      const base64 = btoa(new Uint8Array(arrayBuffer).reduce((s, b) => s + String.fromCharCode(b), ""));
      const res = await fetch("/api/upload", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fileName: file.name, mimeType: file.type, fileData: base64 }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Upload failed.");
      setUrl(data.url);
      if (/\.(mp4|mov|webm|avi)$/i.test(file.name)) setType("video");
    } catch (err) {
      setUploadError(err.message || "Failed to upload file.");
      setUrl("");
    } finally {
      setUploading(false);
    }
  };

  const addResource = () => {
    if (!title.trim() || !url.trim()) return;
    setResources([...resources, { id: `res-${Date.now()}`, title: title.trim(), type, url: url.trim(), fileName: fileName || null }]);
    resetForm();
    showToast(`Added "${title.trim()}"`);
  };
  const removeResource = (id) => setResources(resources.filter((r) => r.id !== id));

  const completionCount = (resourceId) => projectParticipants.filter((p) => (p.completedResourceIds || []).includes(resourceId)).length;

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-bold text-slate-800">Resources — {currentProject?.name}</h2>
          <p className="text-sm text-slate-400">Documents and videos participants review and mark as completed</p>
        </div>
        <button onClick={() => setShowAdd((s) => !s)} className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
          <Plus size={15} /> Add Resource
        </button>
      </div>

      {showAdd && (
        <div className="mb-4 space-y-3 rounded-xl border border-slate-100 bg-slate-50 p-4">
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Title</label>
              <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. Program Handbook" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm" />
            </div>
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Type</label>
              <select value={type} onChange={(e) => setType(e.target.value)} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm">
                <option value="document">Document</option>
                <option value="video">Video</option>
              </select>
            </div>
          </div>

          <div className="flex rounded-lg border border-slate-200 p-1" style={{ maxWidth: 280 }}>
            <button
              type="button"
              onClick={() => setMode("upload")}
              className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-1.5 text-xs font-medium ${mode === "upload" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-white"}`}
            >
              <UploadCloud size={12} /> Upload file
            </button>
            <button
              type="button"
              onClick={() => setMode("link")}
              className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-1.5 text-xs font-medium ${mode === "link" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-white"}`}
            >
              <Link2 size={12} /> Paste a link
            </button>
          </div>

          {mode === "upload" ? (
            <>
              <input
                ref={fileInputRef}
                type="file"
                accept=".doc,.docx,.pdf,.ppt,.pptx,.xls,.xlsx,.mp4,.mov,.webm,.m4v"
                className="hidden"
                onChange={(e) => handleFileUpload(e.target.files && e.target.files[0])}
              />
              <div
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setDragOver(false);
                  const file = e.dataTransfer.files && e.dataTransfer.files[0];
                  if (file) handleFileUpload(file);
                }}
                onClick={() => fileInputRef.current && fileInputRef.current.click()}
                className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-4 py-8 text-center transition-colors ${
                  dragOver ? "border-blue-400 bg-blue-50" : "border-slate-200 bg-white"
                }`}
              >
                {uploading ? (
                  <>
                    <RefreshCw size={20} className="mb-2 animate-spin text-blue-500" />
                    <p className="text-sm font-medium text-slate-600">Uploading {fileName}…</p>
                  </>
                ) : url && fileName ? (
                  <>
                    <CheckCircle2 size={20} className="mb-2 text-emerald-500" />
                    <p className="text-sm font-medium text-slate-600">{fileName}</p>
                    <p className="mt-0.5 text-xs text-slate-400">Uploaded — click to replace</p>
                  </>
                ) : (
                  <>
                    <UploadCloud size={22} className={dragOver ? "text-blue-500" : "text-slate-400"} />
                    <p className="mt-2 text-sm font-medium text-slate-600">Drag &amp; drop a Word, PDF, PowerPoint, or video file, or click to browse</p>
                  </>
                )}
              </div>
              {uploadError && (
                <div className="rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">
                  {uploadError}
                  {uploadError.includes("Blob storage") && (
                    <div className="mt-1">Make sure Blob storage is connected in Vercel → Storage, then redeploy.</div>
                  )}
                </div>
              )}
            </>
          ) : (
            <div>
              <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Link (URL)</label>
              <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://…" className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm" />
              <p className="mt-1 text-xs text-slate-400">Link to a hosted document (Google Drive, SharePoint, etc.) or video (YouTube, Vimeo, Teams recording, etc.)</p>
            </div>
          )}

          <div className="flex justify-end gap-2">
            <button onClick={resetForm} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">Cancel</button>
            <button onClick={addResource} disabled={!title.trim() || !url.trim() || uploading} className="rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40">Add</button>
          </div>
        </div>
      )}

      {resources.length > 0 && (
        <div className="mb-4 flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2" style={{ maxWidth: 320 }}>
          <Search size={15} className="text-slate-400" />
          <input value={resourceSearch} onChange={(e) => setResourceSearch(e.target.value)} placeholder="Search resources…" className="w-full text-sm outline-none placeholder:text-slate-400" />
        </div>
      )}

      {(() => {
        const visibleResources = resources.filter((r) => r.title.toLowerCase().includes(resourceSearch.toLowerCase()));
        if (resources.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 py-14 text-center text-slate-400">
              <FileText size={26} className="mb-2" />
              <p className="text-sm">No resources added yet for this project.</p>
            </div>
          );
        }
        if (visibleResources.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-slate-200 py-14 text-center text-slate-400">
              <Search size={26} className="mb-2" />
              <p className="text-sm">No resources match "{resourceSearch}".</p>
            </div>
          );
        }
        return (
        <div className="space-y-2">
          {visibleResources.map((r) => (
            <div key={r.id} className="flex items-center justify-between gap-3 rounded-xl border border-slate-100 bg-white px-4 py-3">
              <div className="flex items-center gap-3">
                <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-blue-50 text-blue-700">
                  {r.type === "video" ? <FileType size={16} /> : <FileText size={16} />}
                </div>
                <div>
                  <div className="text-sm font-medium text-slate-800">{r.title}</div>
                  {r.fileName ? (
                    <span className="text-xs text-slate-400">{r.fileName}</span>
                  ) : (
                    <a href={r.url} target="_blank" rel="noreferrer" className="text-xs text-blue-600 hover:underline">{r.url}</a>
                  )}
                </div>
              </div>
              <div className="flex items-center gap-3">
                <span className="text-xs text-slate-400">{completionCount(r.id)} / {projectParticipants.length} completed</span>
                <button
                  onClick={() => setReminderResource(r)}
                  disabled={projectParticipants.length === 0}
                  className="flex items-center gap-1 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40"
                >
                  <Mail size={12} /> Remind
                </button>
                <button onClick={() => removeResource(r.id)} className="text-slate-400 hover:text-red-600"><Trash2 size={15} /></button>
              </div>
            </div>
          ))}
        </div>
        );
      })()}
      {reminderResource && (
        <SendResourceReminderModal
          resource={reminderResource}
          projectParticipants={projectParticipants}
          onClose={() => setReminderResource(null)}
        />
      )}
    </div>
  );
}

function GradingCriteriaPanel() {
  const { projects, currentProjectId, updateProject } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId) || projects[0];
  const criteria = currentProject?.criteria || [];
  const [showImport, setShowImport] = useState(false);
  const total = criteria.reduce((s, c) => s + Number(c.weight || 0), 0);
  const balanced = total === 100;

  const setCriteria = (next) => updateProject(currentProject.id, { criteria: next });
  const update = (id, field, value) => setCriteria(criteria.map((c) => (c.id === id ? { ...c, [field]: value } : c)));
  const remove = (id) => setCriteria(criteria.filter((c) => c.id !== id));
  const add = () => setCriteria([...criteria, { id: String(Date.now()), name: "New Criterion", weight: 0, isReflectionCriterion: false }]);
  const setReflectionCriterion = (id) => setCriteria(criteria.map((c) => ({ ...c, isReflectionCriterion: c.id === id })));

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-bold text-slate-800">Grading Criteria — {currentProject?.name}</h2>
          <p className="text-sm text-slate-400">Define criteria and assign weight so it totals 100%. Different projects can have entirely different criteria.</p>
        </div>
        <div className="flex items-center gap-2">
          <span className={`rounded-full px-3 py-1 text-sm font-semibold ${balanced ? "bg-emerald-50 text-emerald-600" : "bg-amber-50 text-amber-600"}`}>
            {balanced ? <CheckCircle2 size={14} className="mr-1 inline" /> : <XCircle size={14} className="mr-1 inline" />}
            Total weight: {total}%
          </span>
          <button onClick={() => setShowImport(true)} className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            <UploadCloud size={14} /> Import from Excel
          </button>
        </div>
      </div>

      <div className="mb-3 flex items-start gap-2 rounded-lg bg-blue-50 px-3 py-2.5 text-xs text-blue-700">
        <Sparkles size={14} className="mt-0.5 shrink-0" />
        Pick one criterion below to be <strong>AI-graded from reflections</strong> (the radio button next to it) — when a participant submits a reflection and it's graded, that score fills in automatically. Every other criterion needs a score entered manually or imported from Excel.
      </div>

      <div className="space-y-2">
        {criteria.map((c) => (
          <div key={c.id} className="flex items-center gap-3 rounded-xl border border-slate-100 bg-slate-50 px-4 py-3">
            <label className="flex shrink-0 items-center gap-1" title="AI-graded from reflections">
              <input type="radio" name="reflectionCriterion" checked={Boolean(c.isReflectionCriterion)} onChange={() => setReflectionCriterion(c.id)} />
              <Sparkles size={13} className={c.isReflectionCriterion ? "text-blue-600" : "text-slate-300"} />
            </label>
            <input
              value={c.name}
              onChange={(e) => update(c.id, "name", e.target.value)}
              className="flex-1 rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-700"
            />
            <div className="flex items-center gap-2">
              <input
                type="number"
                min="0"
                max="100"
                value={c.weight}
                onChange={(e) => update(c.id, "weight", Number(e.target.value))}
                className="w-20 rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-700"
              />
              <span className="text-sm text-slate-400">%</span>
            </div>
            <div className="h-2 w-32 rounded-full bg-slate-200">
              <div className="h-2 rounded-full bg-blue-600" style={{ width: `${Math.min(c.weight, 100)}%` }} />
            </div>
            <button onClick={() => remove(c.id)} className="text-slate-400 hover:text-red-600"><Trash2 size={15} /></button>
          </div>
        ))}
      </div>

      <div className="mt-4 flex items-center justify-between">
        <button onClick={add} className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
          <Plus size={14} /> Add criterion
        </button>
        <span className="text-xs text-slate-400">Changes save automatically</span>
      </div>

      {!balanced && (
        <div className="mt-3 rounded-lg bg-amber-50 px-4 py-2.5 text-xs text-amber-700">
          Weights should add up to 100% before totals can be calculated accurately. Currently at {total}%.
        </div>
      )}

      {showImport && (
        <ImportCriteriaModal
          onClose={() => setShowImport(false)}
          onImport={(newCriteria) => { setCriteria(newCriteria); setShowImport(false); }}
        />
      )}
    </div>
  );
}

function fileIcon(type) {
  if (type === "xlsx") return <FileSpreadsheet size={16} className="text-emerald-600" />;
  if (type === "pdf") return <FileType size={16} className="text-red-500" />;
  if (type === "docx") return <FileText size={16} className="text-blue-600" />;
  return <ImageIcon size={16} className="text-slate-400" />;
}

function FileUploadPanel() {
  const [mode, setMode] = useState("file");
  const [dragOver, setDragOver] = useState(false);

  return (
    <div>
      <div className="mb-4 flex items-center justify-between">
        <div>
          <h2 className="text-lg font-bold text-slate-800">File Upload</h2>
          <p className="text-sm text-slate-400">Import participant data or reflections, or enter details manually</p>
        </div>
        <div className="flex rounded-lg border border-slate-200 p-1">
          <button onClick={() => setMode("file")} className={`rounded-md px-3 py-1.5 text-sm font-medium ${mode === "file" ? "bg-blue-800 text-white" : "text-slate-500"}`}>File upload</button>
          <button onClick={() => setMode("manual")} className={`rounded-md px-3 py-1.5 text-sm font-medium ${mode === "manual" ? "bg-blue-800 text-white" : "text-slate-500"}`}>Manual entry</button>
        </div>
      </div>

      {mode === "file" ? (
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); }}
          className={`flex flex-col items-center justify-center rounded-2xl border-2 border-dashed px-6 py-14 text-center transition-colors ${
            dragOver ? "border-blue-400 bg-blue-50" : "border-slate-200 bg-slate-50"
          }`}
        >
          <UploadCloud size={30} className={dragOver ? "text-blue-500" : "text-slate-400"} />
          <p className="mt-3 text-sm font-medium text-slate-600">Drag & drop Word, Excel, or PDF files here</p>
          <p className="text-xs text-slate-400">or click to browse — participant info can live in the file or the file name</p>
          <button className="mt-4 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">Browse files</button>
          <div className="mt-4 flex gap-4 text-xs text-slate-400">
            <span className="flex items-center gap-1"><FileText size={12} /> .docx</span>
            <span className="flex items-center gap-1"><FileSpreadsheet size={12} /> .xlsx</span>
            <span className="flex items-center gap-1"><FileType size={12} /> .pdf</span>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 gap-4 rounded-2xl border border-slate-100 bg-slate-50 p-6 sm:grid-cols-2">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Full name</label>
            <input className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm" placeholder="e.g. Priya Nair" />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Employee ID</label>
            <input className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm" placeholder="e.g. EMP-5510" />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Email</label>
            <input className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm" placeholder="name@company.com" />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Cohort</label>
            <select className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm">
              <option>Cohort A</option><option>Cohort B</option><option>Cohort C</option>
            </select>
          </div>
          <div className="sm:col-span-2">
            <label className="mb-1 block text-xs font-semibold uppercase text-slate-400">Reflection text</label>
            <textarea rows={4} className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm" placeholder="Paste or type the participant's written reflection…" />
          </div>
          <div className="sm:col-span-2 flex justify-end">
            <button className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900"><Plus size={14} /> Add participant record</button>
          </div>
        </div>
      )}

      <div className="mt-6">
        <h3 className="mb-2 text-sm font-semibold text-slate-700">Recent uploads</h3>
        <div className="overflow-hidden rounded-xl border border-slate-100">
          {recentUploads.map((f, i) => (
            <div key={f.name} className={`flex items-center gap-3 px-4 py-3 text-sm ${i !== recentUploads.length - 1 ? "border-b border-slate-50" : ""}`}>
              {fileIcon(f.type)}
              <div className="flex-1">
                <div className="font-medium text-slate-700">{f.name}</div>
                <div className="text-xs text-slate-400">{f.size} · {f.when}</div>
              </div>
              <span className={`rounded-full px-2 py-0.5 text-xs font-medium ${f.status === "Processed" ? "bg-emerald-50 text-emerald-600" : "bg-sky-50 text-sky-600"}`}>{f.status}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function ScoresFeedbackPanel() {
  const { participants } = useAppData();
  const [q, setQ] = useState("");
  const rows = participants
    .map((p, i) => {
      const a = assessments[i % assessments.length];
      return { ...p, assessment: a.name, score: a.score, status: a.status };
    })
    .filter((r) => (r.name + r.assessment).toLowerCase().includes(q.toLowerCase()));

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-bold text-slate-800">Scores & Feedback</h2>
          <p className="text-sm text-slate-400">Every graded submission across participants</p>
        </div>
        <div className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2">
          <Search size={14} className="text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search participant or assessment…" className="text-sm outline-none placeholder:text-slate-400" />
        </div>
      </div>
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
            <th className="py-2 font-semibold">Participant</th>
            <th className="py-2 font-semibold">Assessment</th>
            <th className="py-2 font-semibold">Score</th>
            <th className="py-2 font-semibold">Status</th>
            <th className="py-2 font-semibold">Feedback</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.employeeId} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
              <td className="py-3">
                <div className="flex items-center gap-2.5">
                  <div className="flex h-7 w-7 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">{initialsFromName(r.name)}</div>
                  <span className="font-medium text-slate-800">{r.name}</span>
                </div>
              </td>
              <td className="py-3 text-slate-500">{r.assessment}</td>
              <td className={`py-3 font-semibold ${r.status === "Pass" ? "text-emerald-600" : r.status === "Fail" ? "text-red-600" : "text-slate-400"}`}>
                {r.score !== null ? `${r.score}%` : "—"}
              </td>
              <td className="py-3"><StatusBadge status={r.status} /></td>
              <td className="py-3">
                <button className="flex items-center gap-1 text-blue-700 hover:underline"><MessageSquare size={13} /> View</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function ThresholdsPanel() {
  const { currentUser } = useAppData();
  const [threshold, setThreshold] = useState(passingThreshold);

  return (
    <div>
      <div className="mb-4">
        <h2 className="text-lg font-bold text-slate-800">Thresholds & Certificates</h2>
        <p className="text-sm text-slate-400">Set the passing score and preview the certificate participants receive</p>
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
        <div className="rounded-2xl border border-slate-100 bg-slate-50 p-6">
          <label className="mb-2 block text-sm font-semibold text-slate-700">Passing threshold</label>
          <div className="flex items-center gap-4">
            <input type="range" min="0" max="100" value={threshold} onChange={(e) => setThreshold(Number(e.target.value))} className="flex-1" />
            <span className="w-16 rounded-lg bg-white px-3 py-1.5 text-center font-bold text-blue-700 border border-slate-200">{threshold}%</span>
          </div>
          <p className="mt-3 text-sm text-slate-500">
            Participants scoring <strong>{threshold}%</strong> or higher automatically receive a certificate of completion.
          </p>
          <div className="mt-4 flex items-center gap-2 rounded-lg bg-white px-4 py-3 text-sm">
            <Users size={15} className="text-blue-600" />
            With this threshold, <strong className="mx-1">{assessments.filter((a) => a.score !== null && a.score >= threshold).length}</strong> of {graded.length} graded submissions would pass.
          </div>
          <button className="mt-4 flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900"><Save size={14} /> Save threshold</button>
        </div>

        <div>
          <div className="mb-2 text-sm font-semibold text-slate-700">Certificate preview</div>
          <div className="overflow-hidden rounded-2xl border-2" style={{ borderColor: BRAND_TEAL }}>
            <div className="relative overflow-hidden p-8 text-center text-white" style={brandGradient()}>
              <LogoWatermark size={160} className="-right-10 -bottom-10" />
              <LogoWatermark size={130} className="-left-8 -top-8" />
              <div className="relative mx-auto mb-2 flex h-11 w-11 items-center justify-center">
                <MercerLogoImg size={40} radius={10} />
              </div>
              <div className="relative text-xs uppercase tracking-widest text-white/70">MercerAssess</div>
              <div className="relative mt-3 text-lg font-bold">Certificate of Completion</div>
              <div className="relative mt-1 text-sm text-white/80">is proudly awarded to</div>
              <div className="relative mt-2 text-xl font-semibold">{currentUser ? currentUser.name : "[Participant Name]"}</div>
              <div className="relative mt-1 text-sm text-white/80">for successfully completing</div>
              <div className="relative text-sm font-medium">{latest.name}</div>
              <div className="relative mt-4 flex items-center justify-center gap-2 text-xs text-white/70">
                <Link2 size={12} /> Verified · Issued {latest.graded}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Leaderboard
// ---------------------------------------------------------------------------

function medalStyle(rank) {
  if (rank === 1) return { bg: "linear-gradient(135deg, #fde68a, #f59e0b)", ring: "#b45309", label: "Gold" };
  if (rank === 2) return { bg: "linear-gradient(135deg, #f1f5f9, #94a3b8)", ring: "#475569", label: "Silver" };
  if (rank === 3) return { bg: "linear-gradient(135deg, #fdba74, #c2703d)", ring: "#9a3412", label: "Bronze" };
  return null;
}

function LeaderboardPanel() {
  const { participants, projects, currentProjectId } = useAppData();
  const currentProject = projects.find((p) => p.id === currentProjectId) || projects[0];
  const projectParticipants = participants
    .filter((p) => (p.projectId || "default") === currentProjectId)
    .map((p) => ({ ...p, displayScore: computeWeightedTotal(p, currentProject) ?? p.avgScore ?? 0 }));
  const ranked = [...projectParticipants].sort((a, b) => (b.displayScore ?? 0) - (a.displayScore ?? 0));
  const podiumOrder = [ranked[1], ranked[0], ranked[2]].filter(Boolean);
  const podiumHeight = { 1: 130, 2: 92, 3: 76 };

  return (
    <div>
      <div className="mb-6">
        <h2 className="text-lg font-bold text-slate-800">Leaderboard</h2>
        <p className="text-sm text-slate-400">Participants ranked by average assessment score</p>
      </div>

      {/* Podium */}
      <div className="mb-8 flex flex-wrap items-end justify-center gap-4">
        {podiumOrder.map((p) => {
          const rank = ranked.indexOf(p) + 1;
          const medal = medalStyle(rank);
          return (
            <div key={p.employeeId} className="flex flex-col items-center" style={{ width: 150 }}>
              <div className="mb-2 flex h-12 w-12 items-center justify-center rounded-full text-white shadow" style={{ backgroundImage: medal.bg }}>
                <Medal size={22} />
              </div>
              <div className="mb-1 flex h-10 w-10 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">
                {initialsFromName(p.name)}
              </div>
              <div className="truncate text-center text-sm font-semibold text-slate-800" style={{ maxWidth: 140 }}>{p.name}</div>
              <div className="text-xs text-slate-400">{p.cohort}</div>
              <div className="mt-1 text-lg font-bold" style={{ color: medal.ring }}>{p.displayScore}%</div>
              <div className="mt-2 flex w-full items-start justify-center rounded-t-lg pt-2" style={{ height: podiumHeight[rank], backgroundImage: medal.bg }}>
                <span className="text-2xl font-bold text-white">{rank}</span>
              </div>
            </div>
          );
        })}
      </div>

      {/* Full ranked list */}
      <div className="overflow-hidden rounded-xl border border-slate-100">
        {ranked.map((p, i) => {
          const rank = i + 1;
          const medal = medalStyle(rank);
          return (
            <div
              key={p.employeeId}
              className={`flex flex-wrap items-center gap-3 px-4 py-3 ${i !== ranked.length - 1 ? "border-b border-slate-50" : ""}`}
              style={medal ? { backgroundColor: rank === 1 ? "#fffbeb" : rank === 2 ? "#f8fafc" : "#fff7ed" } : undefined}
            >
              <div
                className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-sm font-bold"
                style={medal ? { backgroundImage: medal.bg, color: "white" } : { backgroundColor: "#f1f5f9", color: "#64748b" }}
              >
                {medal ? <Medal size={15} /> : rank}
              </div>
              <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">
                {initialsFromName(p.name)}
              </div>
              <div className="min-w-0 flex-1">
                <div className="truncate text-sm font-medium text-slate-800">
                  {p.name}
                  {p.self && <span className="ml-1.5 rounded-full bg-teal-50 px-1.5 py-0.5 font-semibold text-teal-600" style={{ fontSize: 10 }}>YOU</span>}
                  {medal && <span className="ml-1.5 font-medium" style={{ color: medal.ring }}>{medal.label}</span>}
                </div>
                <div className="truncate text-xs text-slate-400">{p.dept} · {p.cohort}</div>
              </div>
              <div className="rounded-full bg-slate-100" style={{ width: 100, height: 8 }}>
                <div
                  className="rounded-full"
                  style={{
                    width: `${Math.min(p.displayScore ?? 0, 100)}%`,
                    height: 8,
                    backgroundColor: (p.displayScore ?? 0) >= passingThreshold ? "#0d9488" : "#f97316",
                  }}
                />
              </div>
              <div className="text-right text-sm font-bold text-slate-800" style={{ width: 48 }}>{p.displayScore}%</div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Admin
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Bulk participant import from an uploaded Excel file — auto-detects which
// sheet and columns hold name/email/ID/cohort/score (real trackers are messy:
// multiple sheets, multi-row headers, inconsistent column names), previews
// what it found, then lets the admin confirm before anything is saved.
// ---------------------------------------------------------------------------

function detectParticipantColumns(headerRow) {
  const norm = (s) => String(s == null ? "" : s).toLowerCase().replace(/\n/g, " ").replace(/\s+/g, " ").trim();
  const headers = headerRow.map(norm);
  // Checks patterns in priority order across ALL headers first (rather than
  // column-by-column), so a more specific pattern like "total out of 100"
  // wins over a column that merely happens to contain "score" earlier in
  // the sheet (e.g. a per-component "Assessment Scores (Weightage: 50%)"
  // column, which is not the same thing as the overall total). Also reports
  // which pattern rank matched (0 = most specific), so a sheet with a
  // precise match can outscore a sheet that only matched loosely, even
  // when both sheets otherwise look equally promising.
  const findCol = (patterns, excludePatterns = []) => {
    for (let rank = 0; rank < patterns.length; rank++) {
      const idx = headers.findIndex((h) => h.includes(patterns[rank]) && !excludePatterns.some((ex) => h.includes(ex)));
      if (idx !== -1) return { index: idx, rank };
    }
    return { index: -1, rank: -1 };
  };
  const scoreMatch = findCol(["total out of 100", "total score", "overall score", "total", "score"], ["weightage"]);
  return {
    name: findCol(["full name", "name"]).index,
    email: findCol(["email"]).index,
    id: findCol(["employee id", "id"], ["email"]).index,
    batch: findCol(["batch", "cohort"]).index,
    dept: findCol(["department"]).index,
    score: scoreMatch.index,
    scoreMatchRank: scoreMatch.rank,
  };
}

function scoreColumnMatch(cols) {
  let score = 0;
  if (cols.name !== -1) score += 2;
  if (cols.email !== -1) score += 2;
  if (cols.id !== -1) score += 1;
  if (cols.batch !== -1) score += 1;
  if (cols.score !== -1) {
    score += 2;
    // Bonus for how specific the score-column match was — a precise "Total
    // out of 100%" (rank 0) beats a loose fallback "score" substring match
    // (rank 4) when comparing sheets that otherwise tie.
    score += Math.max(0, 4 - cols.scoreMatchRank) * 0.1;
  }
  return score;
}

function parseParticipantWorkbook(workbook) {
  let best = null;
  workbook.SheetNames.forEach((sheetName) => {
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    for (let headerRowIdx = 0; headerRowIdx < Math.min(5, rows.length); headerRowIdx++) {
      const cols = detectParticipantColumns(rows[headerRowIdx] || []);
      const s = scoreColumnMatch(cols);
      if (!best || s > best.score) best = { sheetName, headerRowIdx, cols, rows, score: s };
    }
  });
  if (!best || best.score < 5) return { participants: [], sheetName: null, criteriaColumns: [] };

  const { rows, headerRowIdx, cols } = best;
  const headerRow = rows[headerRowIdx] || [];
  const coreIndexes = new Set([cols.name, cols.email, cols.id, cols.batch, cols.dept, cols.score].filter((i) => i !== -1));

  // Every other column with a real header — like the individual milestone,
  // attendance, and coaching columns in a messy real-world tracker — becomes
  // a candidate grading-criteria column. Duplicate headers (e.g. two
  // columns both called "ILM Score") get disambiguated with a suffix.
  const criteriaColumns = [];
  const seenLabels = new Map();
  headerRow.forEach((h, i) => {
    if (coreIndexes.has(i)) return;
    let label = String(h == null ? "" : h).replace(/\n/g, " ").replace(/\s+/g, " ").trim();
    if (!label) return;
    if (seenLabels.has(label)) {
      const n = seenLabels.get(label) + 1;
      seenLabels.set(label, n);
      label = `${label} (${n})`;
    } else {
      seenLabels.set(label, 1);
    }
    criteriaColumns.push({ index: i, label });
  });

  const participants = rows
    .slice(headerRowIdx + 1)
    .map((r) => {
      const name = cols.name !== -1 ? r[cols.name] : null;
      const email = cols.email !== -1 ? r[cols.email] : null;
      if (!name || !email || typeof email !== "string" || !email.includes("@")) return null;
      const rawId = cols.id !== -1 ? r[cols.id] : null;
      const rawBatch = cols.batch !== -1 ? r[cols.batch] : null;
      const rawScore = cols.score !== -1 ? r[cols.score] : null;
      const rawDept = cols.dept !== -1 ? r[cols.dept] : null;
      const parsedScore = typeof rawScore === "number" ? rawScore : parseFloat(rawScore);
      const criteriaValues = {};
      criteriaColumns.forEach((c) => {
        const v = r[c.index];
        if (v !== null && v !== undefined && v !== "") criteriaValues[c.label] = v;
      });
      return {
        employeeId: rawId != null && rawId !== "" ? `EMP-${String(rawId).trim()}` : `EMP-${Math.floor(1000 + Math.random() * 9000)}`,
        name: String(name).trim(),
        email: String(email).trim(),
        cohort: rawBatch != null && rawBatch !== "" ? `Batch ${rawBatch}` : "All Participants",
        dept: rawDept ? String(rawDept).trim() : "",
        avgScore: Number.isFinite(parsedScore) ? Math.round(parsedScore) : 0,
        criteriaValues,
      };
    })
    .filter(Boolean);

  return { participants, sheetName: best.sheetName, criteriaColumns };
}

// ---------------------------------------------------------------------------
// Participant Detail — the full grading picture for one person: every
// criterion in their project, weighted contribution, and the real overall
// total (not just whatever avgScore was last imported).
// ---------------------------------------------------------------------------

function ParticipantDetailModal({ participant, onClose, onEdit }) {
  const { projects } = useAppData();
  const project = projects.find((pr) => pr.id === (participant.projectId || "default")) || projects[0];
  const criteria = project?.criteria || [];
  const weightedTotal = computeWeightedTotal(participant, project);
  const displayTotal = weightedTotal ?? participant.avgScore ?? 0;
  const passed = displayTotal >= passingThreshold;

  const scoredCriteria = criteria.map((c) => {
    const raw = participant.criteriaScores ? participant.criteriaScores[c.name] : undefined;
    const num = typeof raw === "number" ? raw : parseFloat(raw);
    return { ...c, rawValue: raw, numericValue: Number.isFinite(num) ? num : null };
  });

  // Any imported/entered criteria scores that don't match a currently
  // configured criterion (e.g. qualitative tracking columns like "Y"/"not
  // due" from an import) still show up here for reference.
  const otherEntries = Object.entries(participant.criteriaScores || {}).filter(
    ([label]) => !criteria.some((c) => c.name === label)
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-2xl rounded-2xl bg-white p-6 shadow-xl" style={{ maxHeight: "88vh", overflowY: "auto" }}>
        <div className="mb-4 flex items-start justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-full bg-blue-800 text-sm font-semibold text-white">{initialsFromName(participant.name)}</div>
            <div>
              <h3 className="text-lg font-bold text-slate-800">{participant.name}</h3>
              <p className="text-sm text-slate-400">{participant.email} · {participant.cohort} · {participant.dept || "—"}</p>
            </div>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>

        <div className="mb-5 flex flex-wrap items-center gap-4 rounded-xl border border-emerald-100 bg-emerald-50 p-4">
          <div className="text-center" style={{ flex: "1 1 140px" }}>
            <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Overall Score</div>
            <div className="text-4xl font-bold text-emerald-700">{displayTotal.toFixed(1)}</div>
            <div className="text-xs text-slate-400">out of 100 {weightedTotal !== null ? "· weighted by criteria below" : "· from last import (no matching criteria scores yet)"}</div>
          </div>
          <span className="rounded-full bg-emerald-100 px-3 py-1.5 text-xs font-semibold text-emerald-700">{passed ? "✓ PASSED" : "✗ FAILED"}</span>
          <span className="rounded-full bg-white px-3 py-1.5 text-xs font-medium text-slate-500">Threshold: {passingThreshold}</span>
        </div>

        <div className="mb-2 flex items-center gap-2 text-sm font-semibold text-slate-700">
          <SlidersHorizontal size={15} /> Grading Criteria — {project?.name}
        </div>
        {criteria.length === 0 ? (
          <div className="rounded-lg bg-slate-50 px-4 py-6 text-center text-sm text-slate-400">
            This project has no grading criteria configured yet. Set them up in Admin → Grading Criteria.
          </div>
        ) : (
          <div className="overflow-hidden rounded-xl border border-slate-100">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-xs uppercase tracking-wide text-slate-400">
                  <th className="px-4 py-2 font-semibold">Criterion</th>
                  <th className="px-4 py-2 font-semibold">Weight</th>
                  <th className="px-4 py-2 font-semibold">Score</th>
                  <th className="px-4 py-2 font-semibold">Weighted</th>
                </tr>
              </thead>
              <tbody>
                {scoredCriteria.map((c) => (
                  <tr key={c.id} className="border-t border-slate-100">
                    <td className="flex items-center gap-1.5 px-4 py-3 font-medium text-slate-700">
                      {c.isReflectionCriterion && <Sparkles size={12} className="text-blue-600" />}
                      {c.name}
                    </td>
                    <td className="px-4 py-3"><span className="rounded bg-slate-100 px-2 py-0.5 text-xs">{c.weight}%</span></td>
                    <td className="px-4 py-3">
                      {c.numericValue !== null ? (
                        <span className="font-semibold text-teal-700">{c.numericValue}</span>
                      ) : c.rawValue !== undefined ? (
                        <span className="text-xs text-slate-400">{String(c.rawValue)}</span>
                      ) : (
                        <span className="text-xs text-slate-300">No score yet</span>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      {c.numericValue !== null ? (
                        <span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-semibold text-blue-700">+{((c.numericValue * c.weight) / 100).toFixed(1)} pts</span>
                      ) : (
                        <span className="text-xs text-slate-300">—</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {otherEntries.length > 0 && (
          <div className="mt-4">
            <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Other Tracked Data (not part of the weighted total)</div>
            <div className="grid grid-cols-2 gap-1.5" style={{ maxHeight: 160, overflowY: "auto" }}>
              {otherEntries.map(([label, value]) => (
                <div key={label} className="flex items-center justify-between gap-2 rounded bg-slate-50 px-2 py-1.5 text-xs">
                  <span className="truncate text-slate-500" title={label}>{label}</span>
                  <span className="shrink-0 font-medium text-slate-700">{String(value)}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Close
          </button>
          <button
            onClick={() => { onEdit(participant); onClose(); }}
            className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900"
          >
            <Pencil size={14} /> Edit Participant
          </button>
        </div>
      </div>
    </div>
  );
}

function ImportParticipantsModal({ onClose }) {
  const { participants: allParticipants, addParticipant, updateParticipant, showToast, currentProjectId, projects } = useAppData();
  const existingParticipants = allParticipants.filter((p) => (p.projectId || "default") === currentProjectId);
  const currentProject = projects.find((p) => p.id === currentProjectId);
  const [parsing, setParsing] = useState(false);
  const [error, setError] = useState(null);
  const [parsed, setParsed] = useState(null); // { participants, sheetName, criteriaColumns }
  const [included, setIncluded] = useState({}); // employeeId/email -> boolean
  const [criteriaIncluded, setCriteriaIncluded] = useState({}); // label -> boolean
  const [dragOver, setDragOver] = useState(false);
  const [importing, setImporting] = useState(false);
  const fileInputRef = React.useRef(null);

  const handleFile = async (file) => {
    if (!file) return;
    setError(null);
    setParsing(true);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const result = parseParticipantWorkbook(workbook);
      if (!result.sheetName || result.participants.length === 0) {
        setError("Couldn't find recognizable name/email columns in any sheet of this file. Expected columns like Name, Email, ID, Batch, and a total score.");
        setParsed(null);
      } else {
        setParsed(result);
        const initialIncluded = {};
        result.participants.forEach((p) => { initialIncluded[p.email.toLowerCase()] = true; });
        setIncluded(initialIncluded);
        const initialCriteria = {};
        result.criteriaColumns.forEach((c) => { initialCriteria[c.label] = true; });
        setCriteriaIncluded(initialCriteria);
      }
    } catch (err) {
      setError("Couldn't read this file. Make sure it's a valid .xlsx file.");
    } finally {
      setParsing(false);
    }
  };

  const rowsWithStatus = (parsed?.participants || []).map((p) => {
    const existing = existingParticipants.find((ep) => ep.email.toLowerCase() === p.email.toLowerCase());
    return { ...p, isUpdate: Boolean(existing), existingEmployeeId: existing?.employeeId, existingCriteriaScores: existing?.criteriaScores };
  });

  const selectedRows = rowsWithStatus.filter((r) => included[r.email.toLowerCase()]);
  const newCount = selectedRows.filter((r) => !r.isUpdate).length;
  const updateCount = selectedRows.filter((r) => r.isUpdate).length;

  const buildCriteriaScores = (row) => {
    const scores = { ...(row.existingCriteriaScores || {}) };
    Object.entries(row.criteriaValues || {}).forEach(([label, value]) => {
      if (criteriaIncluded[label]) scores[label] = value;
    });
    return scores;
  };

  const handleImport = async () => {
    setImporting(true);
    selectedRows.forEach((row) => {
      const criteriaScores = buildCriteriaScores(row);
      if (row.isUpdate) {
        const updateData = { name: row.name, cohort: row.cohort, avgScore: row.avgScore, criteriaScores };
        if (row.dept) updateData.dept = row.dept;
        updateParticipant(row.existingEmployeeId, updateData);
      } else {
        addParticipant({ name: row.name, email: row.email, cohort: row.cohort, dept: row.dept, avgScore: row.avgScore, criteriaScores });
      }
    });
    showToast(`Imported ${selectedRows.length} participant${selectedRows.length === 1 ? "" : "s"} (${newCount} new, ${updateCount} updated)`);
    setImporting(false);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
      <div className="w-full max-w-2xl rounded-2xl bg-white p-6 shadow-xl" style={{ maxHeight: "85vh", overflowY: "auto" }}>
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-800">Import Participants from Excel</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>
        <div className="mb-4 flex items-center gap-1.5 rounded-lg bg-blue-50 px-3 py-2 text-xs font-medium text-blue-700">
          <Briefcase size={13} /> Importing into: <strong>{currentProject?.name || "Default Project"}</strong> — switch projects first if this isn't the right one.
        </div>

        {!parsed && (
          <>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(e) => handleFile(e.target.files && e.target.files[0])}
            />
            <div
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragOver(false);
                const file = e.dataTransfer.files && e.dataTransfer.files[0];
                if (file) handleFile(file);
              }}
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-4 py-10 text-center transition-colors ${
                dragOver ? "border-blue-400 bg-blue-50" : "border-slate-200 bg-slate-50"
              }`}
            >
              {parsing ? (
                <>
                  <RefreshCw size={22} className="mb-2 animate-spin text-blue-500" />
                  <p className="text-sm font-medium text-slate-600">Reading file…</p>
                </>
              ) : (
                <>
                  <UploadCloud size={24} className={dragOver ? "text-blue-500" : "text-slate-400"} />
                  <p className="mt-2 text-sm font-medium text-slate-600">Drag & drop an .xlsx file, or click to browse</p>
                  <p className="mt-1 text-xs text-slate-400">Auto-detects name, email, employee ID, batch/cohort, and score columns — even across multiple sheets.</p>
                </>
              )}
            </div>
            {error && <div className="mt-3 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{error}</div>}
          </>
        )}

        {parsed && (
          <>
            <div className="mb-3 flex items-center gap-2 rounded-lg bg-blue-50 px-3 py-2.5 text-sm text-blue-700">
              <CheckCircle2 size={15} /> Found {parsed.participants.length} participants in sheet "{parsed.sheetName}"
            </div>
            <div className="mb-3 flex flex-wrap gap-2 text-xs">
              <span className="rounded-full bg-emerald-50 px-3 py-1 font-medium text-emerald-700">{newCount} new</span>
              <span className="rounded-full bg-amber-50 px-3 py-1 font-medium text-amber-700">{updateCount} will be updated (matched by email)</span>
            </div>

            <div className="overflow-hidden rounded-xl border border-slate-100" style={{ maxHeight: 320, overflowY: "auto" }}>
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-slate-50">
                  <tr className="text-left text-xs uppercase tracking-wide text-slate-400">
                    <th className="px-3 py-2"><input type="checkbox" checked={selectedRows.length === rowsWithStatus.length} onChange={(e) => {
                      const next = {};
                      rowsWithStatus.forEach((r) => { next[r.email.toLowerCase()] = e.target.checked; });
                      setIncluded(next);
                    }} /></th>
                    <th className="px-3 py-2 font-semibold">Name</th>
                    <th className="px-3 py-2 font-semibold">Email</th>
                    <th className="px-3 py-2 font-semibold">Cohort</th>
                    <th className="px-3 py-2 font-semibold">Score</th>
                    <th className="px-3 py-2 font-semibold">Status</th>
                  </tr>
                </thead>
                <tbody>
                  {rowsWithStatus.map((r) => (
                    <tr key={r.email} className="border-t border-slate-50">
                      <td className="px-3 py-2">
                        <input
                          type="checkbox"
                          checked={Boolean(included[r.email.toLowerCase()])}
                          onChange={(e) => setIncluded((prev) => ({ ...prev, [r.email.toLowerCase()]: e.target.checked }))}
                        />
                      </td>
                      <td className="px-3 py-2 font-medium text-slate-700">{r.name}</td>
                      <td className="px-3 py-2 text-slate-500">{r.email}</td>
                      <td className="px-3 py-2 text-slate-500">{r.cohort}</td>
                      <td className="px-3 py-2 text-slate-500">{r.avgScore}</td>
                      <td className="px-3 py-2">
                        {r.isUpdate ? (
                          <span className="rounded-full bg-amber-50 px-2 py-0.5 text-xs font-medium text-amber-700">Update</span>
                        ) : (
                          <span className="rounded-full bg-emerald-50 px-2 py-0.5 text-xs font-medium text-emerald-700">New</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {parsed.criteriaColumns.length > 0 && (
              <div className="mt-4 rounded-xl border border-slate-100 bg-slate-50 p-4">
                <div className="mb-2 flex items-center gap-1.5 text-sm font-semibold text-slate-700">
                  <SlidersHorizontal size={14} /> Grading Criteria Columns Found ({parsed.criteriaColumns.length})
                </div>
                <p className="mb-3 text-xs text-slate-400">Every other column in the sheet (attendance, milestones, scores, etc.) is shown below — uncheck any you don't want saved per participant.</p>
                <div className="grid grid-cols-2 gap-2" style={{ maxHeight: 200, overflowY: "auto" }}>
                  {parsed.criteriaColumns.map((c) => (
                    <label key={c.label} className="flex items-center gap-2 rounded-lg bg-white px-3 py-2 text-xs">
                      <input
                        type="checkbox"
                        checked={Boolean(criteriaIncluded[c.label])}
                        onChange={(e) => setCriteriaIncluded((prev) => ({ ...prev, [c.label]: e.target.checked }))}
                      />
                      <span className="truncate text-slate-600" title={c.label}>{c.label}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}

            <button
              onClick={() => { setParsed(null); setError(null); }}
              className="mt-3 text-xs font-medium text-blue-700 hover:underline"
            >
              Choose a different file
            </button>
          </>
        )}

        <div className="mt-6 flex justify-end gap-2">
          <button onClick={onClose} className="rounded-lg border border-slate-200 px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50">
            Cancel
          </button>
          {parsed && (
            <button
              disabled={selectedRows.length === 0 || importing}
              onClick={handleImport}
              className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
            >
              {importing ? <RefreshCw size={14} className="animate-spin" /> : <UploadCloud size={14} />}
              Import {selectedRows.length} Participant{selectedRows.length === 1 ? "" : "s"}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function Admin({ onPreviewAttendance }) {
  const { participants, addParticipant, updateParticipant, deleteParticipant, resetParticipants, projects, currentProjectId, setCurrentProjectId, addProject, deleteProject, updateProject } = useAppData();
  const [tab, setTab] = useState("overview");
  const [query, setQuery] = useState("");
  const [modal, setModal] = useState(null); // null | { mode: "add" } | { mode: "edit", data }
  const [confirmReset, setConfirmReset] = useState(false);
  const [confirmDeleteParticipant, setConfirmDeleteParticipant] = useState(null); // participant object or null
  const [detailParticipant, setDetailParticipant] = useState(null);
  const [showImportModal, setShowImportModal] = useState(false);
  const [showNewProject, setShowNewProject] = useState(false);
  const [confirmDeleteProject, setConfirmDeleteProject] = useState(false);
  const logoInputRef = React.useRef(null);
  const handleLogoUpload = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => updateProject(currentProjectId, { logoUrl: reader.result });
    reader.readAsDataURL(file);
  };
  const [newProjectName, setNewProjectName] = useState("");
  const [cohortFilter, setCohortFilter] = useState("All Cohorts");
  const [sortBy, setSortBy] = useState("name"); // 'name' | 'score-desc' | 'score-asc'
  const currentProject = projects.find((p) => p.id === currentProjectId) || projects[0];
  const projectParticipants = participants.filter((p) => (p.projectId || "default") === currentProjectId);
  const cohortOptions = ["All Cohorts", ...Array.from(new Set(projectParticipants.map((p) => p.cohort).filter(Boolean))).sort()];
  const tabs = [
    { key: "overview", label: "Overview", icon: LayoutGrid },
    { key: "participants", label: "Participants", icon: Users },
    { key: "attendance", label: "Meetings", icon: Calendar },
    { key: "attendanceReport", label: "Attendance Report", icon: UserCheck },
    { key: "leaderboard", label: "Leaderboard", icon: Trophy },
    { key: "groups", label: "Groups", icon: Users },
    { key: "resources", label: "Resources", icon: FileText },
    { key: "criteria", label: "Grading Criteria", icon: SlidersHorizontal },
    { key: "upload", label: "File Upload", icon: Upload },
    { key: "scores", label: "Scores & Feedback", icon: BarChart3 },
    { key: "thresholds", label: "Thresholds & Certs", icon: Award },
    { key: "adminAccounts", label: "Admin Accounts", icon: ShieldCheck },
  ];
  const filtered = projectParticipants
    .filter((p) => (p.name + p.email + p.employeeId).toLowerCase().includes(query.toLowerCase()))
    .filter((p) => cohortFilter === "All Cohorts" || p.cohort === cohortFilter)
    .sort((a, b) => {
      if (sortBy === "score-desc") return (computeWeightedTotal(b, currentProject) ?? b.avgScore ?? 0) - (computeWeightedTotal(a, currentProject) ?? a.avgScore ?? 0);
      if (sortBy === "score-asc") return (computeWeightedTotal(a, currentProject) ?? a.avgScore ?? 0) - (computeWeightedTotal(b, currentProject) ?? b.avgScore ?? 0);
      return a.name.localeCompare(b.name);
    });
  const active = filtered.filter((p) => p.status === "Active").length;

  return (
    <div className="mx-auto max-w-7xl space-y-6 px-6 py-6">
      <div className="relative overflow-hidden rounded-3xl p-8 text-white" style={brandGradient()}>
        <LogoWatermark size={200} className="-right-8 -bottom-12" />
        <div className="flex flex-wrap items-start justify-between gap-6">
          <div className="flex items-start gap-4">
            <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-white/15"><ShieldCheck size={22} /></div>
            <div>
              <span className="rounded bg-white/15 px-2 py-0.5 text-xs font-semibold tracking-wide">ADMIN</span>
              <span className="ml-2 text-sm text-white/70">{adminInfo.email}</span>
              <h1 className="mt-2 text-2xl font-bold">Admin Control Panel</h1>
              <p className="text-sm text-white/70">Manage participants, configure grading, upload data, and issue certificates</p>
            </div>
          </div>
          <div className="flex gap-3">
            <div className="rounded-xl bg-white/10 px-4 py-3 text-center"><div className="text-xs text-white/60">Total Participants</div><div className="text-xl font-bold">{adminInfo.totalParticipants}</div></div>
            <div className="rounded-xl bg-white/10 px-4 py-3 text-center"><div className="text-xs text-white/60">Pending Review</div><div className="text-xl font-bold text-amber-300">{adminInfo.pendingReview}</div></div>
            <div className="rounded-xl bg-white/10 px-4 py-3 text-center"><div className="text-xs text-white/60">Certs Issued</div><div className="text-xl font-bold text-emerald-300">{adminInfo.certsIssued}</div></div>
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white">
        <div className="flex flex-wrap gap-6 border-b border-slate-100 px-6">
          {tabs.map(({ key, label, icon: Icon }) => (
            <button
              key={key}
              onClick={() => setTab(key)}
              className={`flex items-center gap-1.5 border-b-2 py-3 text-sm font-medium ${
                tab === key ? "border-blue-700 text-blue-700" : "border-transparent text-slate-400 hover:text-slate-600"
              }`}
            >
              <Icon size={15} /> {label}
            </button>
          ))}
        </div>

        <div className="p-6">
          {tab === "attendance" && <AttendancePanel onPreviewAttendance={onPreviewAttendance} />}
          {tab === "attendanceReport" && <AttendanceReportPanel />}
          {tab === "leaderboard" && <LeaderboardPanel />}
          {tab === "groups" && <GroupsPanel />}
          {tab === "resources" && <ResourcesPanel />}
          {tab === "criteria" && <GradingCriteriaPanel />}
          {tab === "upload" && <FileUploadPanel />}
          {tab === "scores" && <ScoresFeedbackPanel />}
          {tab === "thresholds" && <ThresholdsPanel />}
          {tab === "adminAccounts" && <AdminAccountsPanel />}
          {tab === "overview" && <OverviewPanel />}
          {tab === "participants" && (
            <>
              <div className="mb-4 rounded-xl border-2 border-blue-100 bg-blue-50 p-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-2">
                    <Briefcase size={18} className="text-blue-700" />
                    <div>
                      <div className="text-xs font-semibold uppercase tracking-wide text-blue-600">Currently viewing project</div>
                      <select
                        value={currentProjectId}
                        onChange={(e) => setCurrentProjectId(e.target.value)}
                        className="mt-0.5 rounded-lg border border-blue-200 bg-white px-3 py-1.5 text-sm font-bold text-slate-800"
                      >
                        {projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                      </select>
                    </div>
                    <div className="ml-2 flex items-center gap-2">
                      <ClientLogo logoUrl={currentProject?.logoUrl} size={40} />
                      <input
                        ref={logoInputRef}
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => handleLogoUpload(e.target.files && e.target.files[0])}
                      />
                      <button
                        onClick={() => logoInputRef.current && logoInputRef.current.click()}
                        className="flex items-center gap-1 rounded-lg border border-blue-200 bg-white px-2.5 py-1.5 text-xs font-medium text-blue-700 hover:bg-blue-100"
                        title="Upload this project's client logo — shown on its surveys and meetings instead of the Mercer logo"
                      >
                        <UploadCloud size={12} /> {currentProject?.logoUrl ? "Change Logo" : "Add Client Logo"}
                      </button>
                    </div>
                  </div>
                  {!showNewProject ? (
                    <button
                      onClick={() => setShowNewProject(true)}
                      className="flex items-center gap-1 rounded-lg border border-blue-200 bg-white px-3 py-2 text-sm font-medium text-blue-700 hover:bg-blue-100"
                    >
                      <Plus size={14} /> New Project
                    </button>
                  ) : (
                    <div className="flex items-center gap-1.5">
                      <input
                        autoFocus
                        value={newProjectName}
                        onChange={(e) => setNewProjectName(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter" && newProjectName.trim()) {
                            addProject(newProjectName.trim());
                            setNewProjectName("");
                            setShowNewProject(false);
                          }
                        }}
                        placeholder="Project name…"
                        className="rounded-lg border border-slate-200 px-3 py-2 text-sm"
                        style={{ width: 160 }}
                      />
                      <button
                        disabled={!newProjectName.trim()}
                        onClick={() => { addProject(newProjectName.trim()); setNewProjectName(""); setShowNewProject(false); }}
                        className="rounded-lg bg-blue-800 px-3 py-2 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
                      >
                        Create
                      </button>
                      <button onClick={() => { setShowNewProject(false); setNewProjectName(""); }} className="text-slate-400 hover:text-slate-600">
                        <X size={16} />
                      </button>
                    </div>
                  )}
                  {currentProjectId !== "default" && !showNewProject && (
                    <button
                      onClick={() => setConfirmDeleteProject(true)}
                      className="flex items-center gap-1 rounded-lg border border-red-200 bg-white px-3 py-2 text-sm font-medium text-red-500 hover:bg-red-50"
                    >
                      <Trash2 size={14} /> Delete Project
                    </button>
                  )}
                </div>
                {projectParticipants.length === 0 && (
                  <div className="mt-3 flex items-start gap-2 rounded-lg bg-white px-3 py-2.5 text-xs text-slate-600">
                    <XCircle size={14} className="mt-0.5 shrink-0 text-amber-500" />
                    <div>
                      <strong>"{currentProject?.name}" has no participants yet</strong> — they haven't been added or imported into this project. Nothing else is missing; switch the dropdown above to see other projects' participants, or use Add Participant / Import from Excel below to populate this one.
                    </div>
                  </div>
                )}
              </div>
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-800">Participants</h2>
                  <p className="text-sm text-slate-400">Manage participants in this project</p>
                </div>
              </div>
              <div className="mb-4 flex flex-wrap items-center gap-3">
                <div className="flex flex-1 items-center gap-2 rounded-lg border border-slate-200 px-3 py-2" style={{ minWidth: 220 }}>
                  <Search size={15} className="text-slate-400" />
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Search by name, email, or ID…"
                    className="w-full text-sm outline-none placeholder:text-slate-400"
                  />
                </div>
                <select value={cohortFilter} onChange={(e) => setCohortFilter(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-600">
                  {cohortOptions.map((c) => <option key={c}>{c}</option>)}
                </select>
                <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-600">
                  <option value="name">Sort: Name</option>
                  <option value="score-desc">Sort: Highest Score</option>
                  <option value="score-asc">Sort: Lowest Score</option>
                </select>
                <button onClick={() => setModal({ mode: "add" })} className="flex items-center gap-1.5 rounded-lg bg-blue-800 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-900">
                  <Plus size={15} /> Add Participant
                </button>
                <button
                  onClick={() => setShowImportModal(true)}
                  className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
                >
                  <UploadCloud size={14} /> Import from Excel
                </button>
                <button
                  onClick={() => downloadRosterCSV(projectParticipants, currentProject)}
                  className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50"
                >
                  <Download size={14} /> Export Roster (CSV)
                </button>
                <button
                  onClick={() => setConfirmReset(true)}
                  className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-500 hover:bg-slate-50"
                  title="Clear saved edits and restore the sample participants"
                >
                  <RefreshCw size={14} /> Reset sample data
                </button>
              </div>

              <div className="mb-4 grid grid-cols-2 gap-4 sm:grid-cols-4">
                <div className="rounded-xl border border-slate-100 bg-slate-50 p-4"><div className="text-xs uppercase text-slate-400">Total</div><div className="text-2xl font-bold text-slate-800">{filtered.length}</div></div>
                <div className="rounded-xl border border-slate-100 bg-slate-50 p-4"><div className="text-xs uppercase text-slate-400">Active</div><div className="text-2xl font-bold text-emerald-600">{active}</div></div>
                <div className="rounded-xl border border-slate-100 bg-slate-50 p-4"><div className="text-xs uppercase text-slate-400">Inactive</div><div className="text-2xl font-bold text-red-500">{filtered.length - active}</div></div>
                <div className="rounded-xl border border-slate-100 bg-slate-50 p-4"><div className="text-xs uppercase text-slate-400">Filtered</div><div className="text-2xl font-bold text-blue-700">{filtered.length}</div></div>
              </div>

              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
                    <th className="py-2 font-semibold">Participant</th>
                    <th className="py-2 font-semibold">Cohort</th>
                    <th className="py-2 font-semibold">Department</th>
                    <th className="py-2 font-semibold">Employee ID</th>
                    <th className="py-2 font-semibold">Assessments</th>
                    <th className="py-2 font-semibold">Last Activity</th>
                    <th className="py-2 font-semibold">Status</th>
                    <th className="py-2 font-semibold">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((p) => (
                    <tr key={p.employeeId} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                      <td className="py-3 cursor-pointer" onClick={() => setDetailParticipant(p)}>
                        <div className="flex items-center gap-2.5">
                          <div className="flex h-8 w-8 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">{initialsFromName(p.name)}</div>
                          <div>
                            <div className="font-medium text-slate-800 hover:text-blue-700">{p.name}{p.self && <span className="ml-1.5 rounded-full bg-teal-50 px-1.5 py-0.5 font-semibold text-teal-600" style={{ fontSize: 10 }}>YOU</span>}</div>
                            <div className="text-xs text-slate-400">{p.email}</div>
                          </div>
                        </div>
                      </td>
                      <td className="py-3"><span className="rounded-full bg-blue-50 px-2 py-0.5 text-xs font-medium text-blue-700">{p.cohort}</span></td>
                      <td className="py-3 text-slate-500">{p.dept}</td>
                      <td className="py-3 text-slate-500">{p.employeeId}</td>
                      <td className="py-3 font-semibold text-slate-700">{p.count}</td>
                      <td className="py-3 text-slate-500">{p.last}</td>
                      <td className="py-3">
                        <span className={`rounded-full px-2 py-0.5 text-xs font-medium ${p.status === "Active" ? "bg-emerald-50 text-emerald-600" : "bg-red-50 text-red-500"}`}>
                          {p.status}
                        </span>
                      </td>
                      <td className="py-3">
                        <div className="flex gap-2 text-slate-400">
                          <button onClick={() => setModal({ mode: "edit", data: p })} className="hover:text-blue-600"><Pencil size={14} /></button>
                          <button
                            onClick={() => setConfirmDeleteParticipant(p)}
                            className="hover:text-red-600"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {filtered.length === 0 && (
                <div className="flex flex-col items-center justify-center py-14 text-center text-slate-400">
                  <Users size={26} className="mb-2" />
                  <p className="text-sm">No participants match "{query}".</p>
                </div>
              )}
            </>
          )}
        </div>
      </div>

      {modal && (
        <ParticipantFormModal
          title={modal.mode === "add" ? "Add Participant" : "Edit Participant"}
          initial={modal.mode === "edit" ? modal.data : null}
          showStatus
          onClose={() => setModal(null)}
          onSave={(data) => {
            if (modal.mode === "add") addParticipant(data);
            else updateParticipant(modal.data.employeeId, data);
            setModal(null);
          }}
        />
      )}
      {showImportModal && <ImportParticipantsModal onClose={() => setShowImportModal(false)} />}
      {detailParticipant && (
        <ParticipantDetailModal
          participant={detailParticipant}
          onClose={() => setDetailParticipant(null)}
          onEdit={(p) => setModal({ mode: "edit", data: p })}
        />
      )}
      {confirmDeleteProject && (
        <ConfirmModal
          title={`Delete "${currentProject?.name}"?`}
          message={`Its ${projectParticipants.length} participant${projectParticipants.length === 1 ? "" : "s"} will be moved to Default Project rather than deleted — no participant data is lost, just this project grouping. This can't be undone.`}
          confirmLabel="Delete Project"
          onCancel={() => setConfirmDeleteProject(false)}
          onConfirm={() => { deleteProject(currentProjectId); setConfirmDeleteProject(false); }}
        />
      )}
      {confirmReset && (
        <ConfirmModal
          title="Reset all participants?"
          message="This clears any edits you've made and restores the built-in sample data. This can't be undone."
          confirmLabel="Reset"
          onCancel={() => setConfirmReset(false)}
          onConfirm={() => {
            resetParticipants();
            setConfirmReset(false);
          }}
        />
      )}
      {confirmDeleteParticipant && (
        <ConfirmModal
          title="Remove this participant?"
          message={`${confirmDeleteParticipant.name} will be permanently removed from the participant list.`}
          onCancel={() => setConfirmDeleteParticipant(null)}
          onConfirm={() => {
            deleteParticipant(confirmDeleteParticipant.employeeId, confirmDeleteParticipant.name);
            setConfirmDeleteParticipant(null);
          }}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Login
// ---------------------------------------------------------------------------

const featureBullets = [
  { icon: FileText, text: "AI-powered reflection grading with detailed feedback" },
  { icon: BarChart3, text: "Per-criterion score breakdown with weighted analysis" },
  { icon: Award, text: "Instant certificate generation upon passing" },
  { icon: PieChartIcon, text: "Survey analytics and cohort performance insights" },
];

const platformStats = [
  { value: 2400, suffix: "+", label: "Participants graded" },
  { value: 94, suffix: "%", label: "Satisfaction rate" },
  { value: 38, suffix: "", label: "Programs active" },
];

function CountUp({ to, duration = 1400, suffix = "" }) {
  const [value, setValue] = useState(0);
  useEffect(() => {
    let raf;
    let start = null;
    const step = (timestamp) => {
      if (start === null) start = timestamp;
      const progress = Math.min((timestamp - start) / duration, 1);
      // ease-out for a natural "settling" feel rather than linear ticking
      const eased = 1 - Math.pow(1 - progress, 3);
      setValue(Math.round(eased * to));
      if (progress < 1) raf = requestAnimationFrame(step);
    };
    raf = requestAnimationFrame(step);
    return () => cancelAnimationFrame(raf);
  }, [to, duration]);
  return <>{value.toLocaleString()}{suffix}</>;
}

function LoginPage({ onParticipantLogin, onAdminLogin }) {
  const { signup, verifyLogin, requestPasswordReset, verifyAdminLogin, requestAdminPasswordReset } = useAppData();
  const [role, setRole] = useState("participant");
  const [mode, setMode] = useState("signin"); // 'signin' | 'signup' | 'forgot', participant tab only
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [remember, setRemember] = useState(false);
  const [error, setError] = useState(null);
  const [submitting, setSubmitting] = useState(false);
  const [forgotEmail, setForgotEmail] = useState("");
  const [forgotSubmitting, setForgotSubmitting] = useState(false);
  const [forgotMessage, setForgotMessage] = useState(null);

  const demoAccounts = [
    { role: "participant", label: "Participant", email: "Mokhtar.Alkhlifa@acmecorp.com", password: "demo1234" },
    { role: "admin", label: "Admin", email: adminInfo.email, password: "demo1234" },
  ];

  const useDemo = (acct) => {
    setRole(acct.role);
    setMode("signin");
    setEmail(acct.email);
    setPassword(acct.password);
    setError(null);
  };

  const switchMode = (next) => {
    setMode(next);
    setError(null);
    setName("");
    setConfirmPassword("");
  };

  const handleForgotSubmit = async () => {
    if (!forgotEmail.trim()) return;
    setForgotSubmitting(true);
    setForgotMessage(null);
    try {
      const result = role === "admin" ? await requestAdminPasswordReset(forgotEmail) : await requestPasswordReset(forgotEmail);
      setForgotMessage(result.message);
    } catch (err) {
      setForgotMessage("If an account exists for that email, a reset link has been sent.");
    } finally {
      setForgotSubmitting(false);
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError(null);

    if (role === "admin") {
      setSubmitting(true);
      try {
        const result = await verifyAdminLogin(email, password);
        if (!result.ok) return setError(result.error);
        onAdminLogin();
      } catch (err) {
        setError("Something went wrong. Please try again.");
      } finally {
        setSubmitting(false);
      }
      return;
    }

    if (mode === "signup") {
      if (!name.trim()) return setError("Enter your name.");
      if (password.length < 6) return setError("Password must be at least 6 characters.");
      if (password !== confirmPassword) return setError("Passwords don't match.");
      setSubmitting(true);
      try {
        const result = await signup(name, email, password);
        if (!result.ok) return setError(result.error);
        onParticipantLogin();
      } catch (err) {
        setError("Something went wrong. Please try again.");
      } finally {
        setSubmitting(false);
      }
      return;
    }

    setSubmitting(true);
    try {
      const result = await verifyLogin(email, password, remember);
      if (!result.ok) return setError(result.error);
      onParticipantLogin();
    } catch (err) {
      setError("Something went wrong. Please try again.");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="flex min-h-screen flex-col md:flex-row">
      {/* Left brand panel */}
      <div className="relative flex w-full flex-col justify-between overflow-hidden p-10 text-white md:w-1/2" style={brandGradient("135deg")}>
        <style>{`
          @keyframes mercerFadeUp {
            0% { opacity: 0; transform: translateY(14px); }
            100% { opacity: 1; transform: translateY(0); }
          }
          @keyframes mercerDrift1 {
            0%, 100% { transform: translate(0, 0) scale(1); }
            50% { transform: translate(-16px, 18px) scale(1.08); }
          }
          @keyframes mercerDrift2 {
            0%, 100% { transform: translate(0, 0) scale(1); }
            50% { transform: translate(14px, -14px) scale(1.1); }
          }
          @keyframes mercerWatermarkBreathe {
            0%, 100% { opacity: 0.08; transform: scale(1) rotate(0deg); }
            50% { opacity: 0.2; transform: scale(1.06) rotate(2deg); }
          }
          @keyframes mercerWatermarkBreatheSlow {
            0%, 100% { opacity: 0.1; transform: scale(1); }
            50% { opacity: 0.22; transform: scale(1.08); }
          }
        `}</style>

        <LogoWatermark
          size={120}
          className="-right-6 -top-10"
          style={{ animation: "mercerDrift1 9s ease-in-out infinite, mercerWatermarkBreathe 5s ease-in-out infinite" }}
        />
        <LogoWatermark
          size={90}
          className="right-40"
          style={{ bottom: -50, animation: "mercerDrift2 11s ease-in-out infinite, mercerWatermarkBreathe 6s ease-in-out infinite 1.2s" }}
        />
        <LogoWatermark
          size={260}
          className="-right-16"
          style={{ bottom: -80, animation: "mercerWatermarkBreatheSlow 7s ease-in-out infinite 0.4s" }}
        />

        <div className="relative" style={{ animation: "mercerFadeUp 600ms ease-out both" }}>
          <Logo light size={40} animated />
          <div className="mt-1 pl-1 text-xs text-white/60">a business of Marsh McLennan</div>
        </div>

        <div className="relative max-w-lg">
          <div
            className="mb-5 inline-flex items-center gap-1.5 rounded-full bg-white/15 px-4 py-1.5 text-xs font-medium"
            style={{ animation: "mercerFadeUp 600ms ease-out 80ms both" }}
          >
            <Medal size={13} /> Assessment &amp; Certification Platform
          </div>
          <h1 className="text-4xl font-bold leading-tight" style={{ animation: "mercerFadeUp 600ms ease-out 160ms both" }}>
            Measure growth.<br />Recognize excellence.
          </h1>
          <p className="mt-4 text-white/70" style={{ animation: "mercerFadeUp 600ms ease-out 240ms both" }}>
            MercerAssess provides a rigorous, transparent assessment experience — from reflection grading to certified achievement. Your progress, clearly measured.
          </p>
          <ul className="mt-8 space-y-3">
            {featureBullets.map(({ icon: Icon, text }, i) => (
              <li
                key={text}
                className="flex items-center gap-3 text-sm text-white/85"
                style={{ animation: `mercerFadeUp 600ms ease-out ${320 + i * 90}ms both` }}
              >
                <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-md bg-white/15"><Icon size={13} /></span>
                {text}
              </li>
            ))}
          </ul>
        </div>

        <div className="relative flex flex-wrap gap-4">
          {platformStats.map((s, i) => (
            <div
              key={s.label}
              className="rounded-xl bg-white/10 px-5 py-3"
              style={{ flex: "1 1 140px", animation: `mercerFadeUp 600ms ease-out ${680 + i * 90}ms both` }}
            >
              <div className="text-xl font-bold"><CountUp to={s.value} suffix={s.suffix} duration={1400 + i * 200} /></div>
              <div className="text-xs text-white/70">{s.label}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Right sign-in panel */}
      <div className="flex w-full flex-1 items-center justify-center bg-slate-50 p-8 md:w-1/2">
        <div className="w-full" style={{ maxWidth: 420 }}>
          <h2 className="text-2xl font-bold text-slate-800">{role === "admin" ? "Admin sign in" : mode === "signup" ? "Create your account" : "Sign in"}</h2>
          {role === "participant" && (
            <p className="mb-6 text-sm text-slate-500">
              {mode === "signup" ? (
                <>Already have an account? <button type="button" onClick={() => switchMode("signin")} className="font-semibold text-blue-700 hover:underline">Sign in</button></>
              ) : (
                <>Don't have an account? <button type="button" onClick={() => switchMode("signup")} className="font-semibold text-blue-700 hover:underline">Create one</button></>
              )}
            </p>
          )}
          {role === "admin" && <p className="mb-6 text-sm text-slate-500">Internal admin access.</p>}

          <div className="mb-5 flex rounded-lg border border-slate-200 bg-white p-1">
            <button
              type="button"
              onClick={() => { setRole("participant"); setError(null); }}
              className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-2 text-sm font-medium ${role === "participant" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-slate-50"}`}
            >
              <User size={14} /> Participant
            </button>
            <button
              type="button"
              onClick={() => { setRole("admin"); setError(null); }}
              className={`flex flex-1 items-center justify-center gap-1.5 rounded-md py-2 text-sm font-medium ${role === "admin" ? "bg-blue-800 text-white" : "text-slate-500 hover:bg-slate-50"}`}
            >
              <ShieldCheck size={14} /> Admin
            </button>
          </div>

          <div className="space-y-4">
            {mode === "forgot" ? (
              <>
                <p className="text-sm text-slate-500">Enter your email and we'll send a link to reset your password.</p>
                <div>
                  <label className="mb-1 block text-sm font-medium text-slate-700">Email address</label>
                  <input
                    autoFocus
                    type="text"
                    value={forgotEmail}
                    onChange={(e) => setForgotEmail(e.target.value)}
                    placeholder="you@organization.com"
                    className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-blue-400"
                  />
                </div>
                {forgotMessage && (
                  <div className="rounded-lg bg-emerald-50 px-3 py-2.5 text-xs text-emerald-700">{forgotMessage}</div>
                )}
                <button
                  type="button"
                  onClick={handleForgotSubmit}
                  disabled={forgotSubmitting || !forgotEmail.trim()}
                  className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-800 py-3 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-60"
                >
                  {forgotSubmitting ? <RefreshCw size={15} className="animate-spin" /> : <Mail size={15} />}
                  {forgotSubmitting ? "Sending…" : "Send Reset Link"}
                </button>
                <button
                  type="button"
                  onClick={() => { setMode("signin"); setForgotMessage(null); }}
                  className="w-full text-center text-sm font-medium text-blue-700 hover:underline"
                >
                  Back to sign in
                </button>
              </>
            ) : (
              <>
            {role === "participant" && mode === "signup" && (
              <div>
                <label className="mb-1 block text-sm font-medium text-slate-700">Full name</label>
                <input
                  autoFocus
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  placeholder="Your full name"
                  className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-blue-400"
                />
              </div>
            )}
            <div>
              <label className="mb-1 block text-sm font-medium text-slate-700">Email address</label>
              <input
                type="text"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="you@organization.com"
                className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-blue-400"
              />
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-slate-700">Password</label>
              <div className="relative">
                <input
                  type={showPassword ? "text" : "password"}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  className="w-full rounded-lg border border-slate-200 px-3 py-2.5 pr-10 text-sm outline-none focus:border-blue-400"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword((s) => !s)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            {role === "participant" && mode === "signup" && (
              <div>
                <label className="mb-1 block text-sm font-medium text-slate-700">Confirm password</label>
                <input
                  type={showPassword ? "text" : "password"}
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                  placeholder="••••••••"
                  className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-blue-400"
                />
              </div>
            )}

            {(role === "participant" || role === "admin") && mode === "signin" && (
              <div className="flex items-center justify-between text-sm">
                <label className="flex items-center gap-2 text-slate-600">
                  <input type="checkbox" checked={remember} onChange={(e) => setRemember(e.target.checked)} className="rounded border-slate-300" />
                  Remember me
                </label>
                <button type="button" onClick={() => { setMode("forgot"); setError(null); setForgotEmail(email); }} className="font-medium text-blue-700 hover:underline">Forgot password?</button>
              </div>
            )}

            {error && (
              <div className="rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{error}</div>
            )}

            <button
              type="button"
              onClick={handleSubmit}
              disabled={submitting}
              className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-800 py-3 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-60"
            >
              {submitting ? <RefreshCw size={15} className="animate-spin" /> : <LogOut size={15} style={{ transform: "scaleX(-1)" }} />}
              {submitting
                ? "Please wait…"
                : role === "admin"
                ? "Sign In as Admin"
                : mode === "signup"
                ? "Create Account"
                : "Sign In as Participant"}
            </button>
              </>
            )}
          </div>

          <div className="mt-6 rounded-xl border border-blue-100 bg-blue-50 p-4">
            <div className="mb-3 flex items-center gap-1.5 text-xs font-semibold text-blue-700">
              <MousePointerClick size={13} /> Demo Accounts — Click to autofill
            </div>
            <div className="space-y-2">
              {demoAccounts.map((acct) => (
                <div key={acct.role} className="flex items-center justify-between rounded-lg bg-white px-3 py-2.5">
                  <div>
                    <div className="text-sm font-semibold text-slate-800">{acct.label}</div>
                    <div className="text-xs text-slate-400">{acct.email}</div>
                  </div>
                  <button
                    type="button"
                    onClick={() => useDemo(acct)}
                    className="flex items-center gap-1 text-sm font-medium text-blue-700 hover:underline"
                  >
                    <MousePointerClick size={13} /> Use
                  </button>
                </div>
              ))}
            </div>
            <p className="mt-3 text-xs text-blue-700">Anyone else can click "Create one" above to sign up — with their own email, whether or not an admin has added them yet.</p>
          </div>

          <p className="mt-6 text-center text-xs text-slate-400">
            © {new Date().getFullYear()} Mercer, a business of Marsh McLennan. All rights reserved. · Confidential platform.
          </p>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Root
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Public survey-taking page — reached via QR code or shared link
// (?survey=<id>), with no login required. Fetches the survey's question
// definitions from shared storage, lets the person answer, and posts their
// response to /api/survey-response, which is what powers the live analysis
// back in Survey Analytics.
// ---------------------------------------------------------------------------

// A labeled 5-point rating scale (Strongly Disagree → Strongly Agree),
// shared between the anonymous QR survey form and the in-app participant
// survey flow so both show the same clear labels instead of bare numbers.
function RatingScaleInput({ value, onChange }) {
  const colors = {
    5: { badge: "#059669", bg: "#ecfdf5", border: "#a7f3d0", selectedBg: "#a7f3d0", selectedBorder: "#059669", selectedText: "#047857" },
    4: { badge: "#0d9488", bg: "#f0fdfa", border: "#99f6e4", selectedBg: "#99f6e4", selectedBorder: "#0d9488", selectedText: "#0f766e" },
    3: { badge: "#64748b", bg: "#f8fafc", border: "#e2e8f0", selectedBg: "#e2e8f0", selectedBorder: "#64748b", selectedText: "#475569" },
    2: { badge: "#f97316", bg: "#fff7ed", border: "#fed7aa", selectedBg: "#fed7aa", selectedBorder: "#f97316", selectedText: "#c2410c" },
    1: { badge: "#dc2626", bg: "#fef2f2", border: "#fecaca", selectedBg: "#fecaca", selectedBorder: "#dc2626", selectedText: "#b91c1c" },
  };
  return (
    <div className="space-y-2">
      {[5, 4, 3, 2, 1].map((n) => {
        const selected = value === n;
        const c = colors[n];
        return (
          <button
            key={n}
            type="button"
            onClick={() => onChange(n)}
            className="flex w-full items-center gap-3 rounded-lg border-2 px-3 py-2.5 text-left text-sm transition-colors hover:opacity-90"
            style={{
              borderColor: selected ? c.selectedBorder : c.border,
              backgroundColor: selected ? c.selectedBg : c.bg,
            }}
          >
            <span
              className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-bold text-white"
              style={{ backgroundColor: c.badge }}
            >
              {n}
            </span>
            <span style={{ color: selected ? c.selectedText : "#334155", fontWeight: selected ? 700 : 500 }}>{LIKERT_LABELS_5[n - 1]}</span>
          </button>
        );
      })}
    </div>
  );
}

function TakeSurveyPage({ surveyId }) {
  const [status, setStatus] = useState("loading"); // loading | ready | notfound | submitting | done | error
  const [survey, setSurvey] = useState(null);
  const [answers, setAnswers] = useState({});

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch(`/api/data?key=${encodeURIComponent(SURVEYS_STORAGE_KEY)}`);
        if (!res.ok) throw new Error("not found");
        const data = await res.json();
        const list = JSON.parse(data.value || "[]");
        const found = list.find((s) => s.id === surveyId && s.sourceType === "built");
        if (!found) { setStatus("notfound"); return; }
        setSurvey(found);
        setStatus("ready");
      } catch (err) {
        setStatus("notfound");
      }
    })();
  }, [surveyId]);

  const setAnswer = (qid, value) => setAnswers((a) => ({ ...a, [qid]: value }));
  const canSubmit = survey && survey.questions.every((q) => {
    const v = answers[q.id];
    return q.type === "rating" ? typeof v === "number" : true; // open-text questions are optional
  });

  const handleSubmit = async () => {
    setStatus("submitting");
    try {
      const res = await fetch(`/api/survey-response?surveyId=${encodeURIComponent(surveyId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ answers }),
      });
      if (!res.ok) throw new Error("failed");
      setStatus("done");
    } catch (err) {
      setStatus("error");
    }
  };

  return (
    <div className="relative flex min-h-screen items-center justify-center overflow-hidden p-6" style={brandGradient()}>
      <LogoWatermark size={280} className="-right-16" style={{ bottom: -80 }} />
      <div className="relative w-full" style={{ maxWidth: 560 }}>
        <div className="mb-6 flex justify-center">{survey?.logoUrl ? <ClientLogo logoUrl={survey.logoUrl} size={44} radius={10} /> : <Logo light size={36} />}</div>

        {status === "loading" && (
          <div className="flex flex-col items-center rounded-2xl border border-slate-200 bg-white p-10 text-center text-slate-400">
            <RefreshCw size={22} className="mb-2 animate-spin text-blue-500" />
            Loading survey…
          </div>
        )}

        {status === "notfound" && (
          <div className="flex flex-col items-center rounded-2xl border border-slate-200 bg-white p-10 text-center text-slate-400">
            <PieChartIcon size={26} className="mb-2" />
            <p className="text-sm">This survey link isn't valid, or the survey has been removed.</p>
          </div>
        )}

        {status === "done" && (
          <div className="flex flex-col items-center rounded-2xl border border-slate-200 bg-white p-10 text-center">
            <CheckCircle2 size={30} className="mb-3 text-emerald-500" />
            <h2 className="text-lg font-bold text-slate-800">Thanks for your feedback!</h2>
            <p className="mt-1 text-sm text-slate-500">Your response has been recorded.</p>
          </div>
        )}

        {(status === "ready" || status === "submitting" || status === "error") && survey && (
          <div className="rounded-2xl border border-slate-200 bg-white p-6">
            <h2 className="text-lg font-bold text-slate-800">{survey.name}</h2>
            <p className="mb-5 text-sm text-slate-400">Your answers are anonymous and help improve future programs.</p>

            <div className="space-y-5">
              {survey.questions.map((q, i) => (
                <div key={q.id}>
                  <label className="mb-2 block text-sm font-medium text-slate-700">{i + 1}. {q.text}</label>
                  {q.type === "rating" ? (
                    <RatingScaleInput value={answers[q.id]} onChange={(n) => setAnswer(q.id, n)} />
                  ) : (
                    <textarea
                      rows={3}
                      value={answers[q.id] || ""}
                      onChange={(e) => setAnswer(q.id, e.target.value)}
                      placeholder="Your answer…"
                      className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-blue-400"
                    />
                  )}
                </div>
              ))}
            </div>

            {status === "error" && (
              <div className="mt-4 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">
                Something went wrong submitting your response. Please try again.
              </div>
            )}

            <button
              disabled={!canSubmit || status === "submitting"}
              onClick={handleSubmit}
              className="mt-6 flex w-full items-center justify-center gap-2 rounded-lg bg-blue-800 py-3 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
            >
              {status === "submitting" ? <RefreshCw size={15} className="animate-spin" /> : <CheckCircle2 size={15} />}
              Submit
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Public attendance check-in page — reached via QR code or shared link
// (?attendance=<id>), with no login required. Just a name (and optional
// employee ID), one tap to check in.
// ---------------------------------------------------------------------------

function TakeAttendancePage({ sessionId, onExit }) {
  const { currentUser, verifyLogin } = useAppData();
  const [status, setStatus] = useState("loading"); // loading | ready | notfound | submitting | done | error
  const [session, setSession] = useState(null);
  const [alreadyCheckedIn, setAlreadyCheckedIn] = useState(false);
  const [checkingDuplicate, setCheckingDuplicate] = useState(true);

  // Only used while the visitor isn't signed in yet.
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loginError, setLoginError] = useState(null);
  const [loggingIn, setLoggingIn] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch(`/api/data?key=${encodeURIComponent(ATTENDANCE_STORAGE_KEY)}`);
        if (!res.ok) throw new Error("not found");
        const data = await res.json();
        const list = JSON.parse(data.value || "[]");
        const found = list.find((s) => s.id === sessionId);
        if (!found) { setStatus("notfound"); return; }
        setSession(found);
        setStatus("ready");
      } catch (err) {
        setStatus("notfound");
      }
    })();
  }, [sessionId]);

  useEffect(() => {
    if (!currentUser) { setCheckingDuplicate(false); return; }
    let cancelled = false;
    setCheckingDuplicate(true);
    (async () => {
      try {
        const res = await fetch(`/api/attendance?sessionId=${encodeURIComponent(sessionId)}`);
        const data = res.ok ? await res.json() : { checkins: [] };
        const found = (data.checkins || []).some((c) => c.employeeId === currentUser.employeeId);
        if (!cancelled) setAlreadyCheckedIn(found);
      } catch (err) {
        if (!cancelled) setAlreadyCheckedIn(false);
      } finally {
        if (!cancelled) setCheckingDuplicate(false);
      }
    })();
    return () => { cancelled = true; };
  }, [currentUser, sessionId]);

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoggingIn(true);
    setLoginError(null);
    try {
      const result = await verifyLogin(email, password);
      if (!result.ok) setLoginError(result.error);
    } catch (err) {
      setLoginError("Something went wrong. Please try again.");
    } finally {
      setLoggingIn(false);
    }
  };

  const handleCheckIn = async () => {
    setStatus("submitting");
    try {
      const res = await fetch(`/api/attendance?sessionId=${encodeURIComponent(sessionId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: currentUser.name, employeeId: currentUser.employeeId, cohort: currentUser.cohort, dept: currentUser.dept }),
      });
      if (!res.ok) throw new Error("failed");
      setStatus("done");
    } catch (err) {
      setStatus("error");
    }
  };

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-50 p-6">
      <div className="w-full" style={{ maxWidth: 440 }}>
        {onExit && (
          <button onClick={onExit} className="mb-4 flex items-center gap-1 text-sm font-medium text-slate-500 hover:text-slate-700">
            <ChevronRight size={14} style={{ transform: "scaleX(-1)" }} /> Back to admin (preview mode)
          </button>
        )}
        <div className="mb-6 flex justify-center">{session?.logoUrl ? <ClientLogo logoUrl={session.logoUrl} size={44} radius={10} /> : <Logo size={36} />}</div>

        {status === "loading" && (
          <div className="flex flex-col items-center rounded-2xl border border-slate-200 bg-white p-10 text-center text-slate-400">
            <RefreshCw size={22} className="mb-2 animate-spin text-blue-500" />
            Loading session…
          </div>
        )}

        {status === "notfound" && (
          <div className="flex flex-col items-center rounded-2xl border border-slate-200 bg-white p-10 text-center text-slate-400">
            <UserCheck size={26} className="mb-2" />
            <p className="text-sm">This check-in link isn't valid, or the session has been removed.</p>
          </div>
        )}

        {status === "done" && (
          <div className="flex flex-col items-center rounded-2xl border border-slate-200 bg-white p-10 text-center">
            <CheckCircle2 size={30} className="mb-3 text-emerald-500" />
            <h2 className="text-lg font-bold text-slate-800">You're checked in!</h2>
            <p className="mt-1 text-sm text-slate-500">{session?.name}</p>
          </div>
        )}

        {(status === "ready" || status === "submitting" || status === "error") && session && (
          <div className="rounded-2xl border border-slate-200 bg-white p-6">
            <div className="mb-1 flex items-center gap-2">
              <UserCheck size={18} className="text-blue-700" />
              <h2 className="text-lg font-bold text-slate-800">Check in</h2>
            </div>
            <p className="mb-5 text-sm text-slate-400">{session.name} · {session.date}{session.startTime ? ` · ${session.startTime}` : ""}</p>

            {!currentUser ? (
              <>
                <p className="mb-4 rounded-lg bg-blue-50 px-3 py-2.5 text-xs text-blue-700">
                  Sign in with your own account to check in — attendance is only ever recorded for the person actually signed in, so no one can check in on someone else's behalf.
                </p>
                <form onSubmit={handleLogin} className="space-y-3">
                  <div>
                    <label className="mb-1 block text-sm font-medium text-slate-700">Email address</label>
                    <input
                      type="email"
                      required
                      value={email}
                      onChange={(e) => setEmail(e.target.value)}
                      placeholder="you@organization.com"
                      className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-blue-400"
                    />
                  </div>
                  <div>
                    <label className="mb-1 block text-sm font-medium text-slate-700">Password</label>
                    <div className="relative">
                      <input
                        type={showPassword ? "text" : "password"}
                        required
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                        className="w-full rounded-lg border border-slate-200 px-3 py-2.5 pr-10 text-sm outline-none focus:border-blue-400"
                      />
                      <button type="button" onClick={() => setShowPassword((s) => !s)} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400">
                        {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                      </button>
                    </div>
                  </div>
                  {loginError && <div className="rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{loginError}</div>}
                  <button
                    type="submit"
                    disabled={loggingIn}
                    className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-800 py-3 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    {loggingIn ? <RefreshCw size={15} className="animate-spin" /> : <UserCheck size={15} />}
                    Sign In &amp; Check In
                  </button>
                </form>
                <p className="mt-3 text-center text-xs text-slate-400">Don't have an account yet? Create one in the main app first, then come back to this link.</p>
              </>
            ) : (
              <>
                <div className="mb-4 flex items-center gap-3 rounded-lg bg-slate-50 p-3">
                  <div className="flex h-9 w-9 items-center justify-center rounded-full bg-blue-800 text-xs font-semibold text-white">{initialsFromName(currentUser.name)}</div>
                  <div>
                    <div className="text-sm font-medium text-slate-800">{currentUser.name}</div>
                    <div className="text-xs text-slate-400">{currentUser.cohort}</div>
                  </div>
                </div>
                {checkingDuplicate ? (
                  <div className="mb-2 flex justify-center py-1 text-slate-400"><RefreshCw size={16} className="animate-spin" /></div>
                ) : alreadyCheckedIn ? (
                  <p className="mb-4 flex items-center gap-1 text-xs font-medium text-amber-600">
                    <CheckCircle2 size={12} /> You're already checked in for this meeting.
                  </p>
                ) : null}
                {currentUser.emailVerified === false && (
                  <div className="mb-4 rounded-lg bg-amber-50 px-3 py-2.5 text-xs text-amber-700">
                    Please verify your email before checking in — check your inbox for the verification link from signup, or resend it from your Dashboard.
                  </div>
                )}
                {status === "error" && (
                  <div className="mb-4 rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">Something went wrong checking you in. Please try again.</div>
                )}
                <button
                  disabled={alreadyCheckedIn || checkingDuplicate || status === "submitting" || currentUser.emailVerified === false}
                  onClick={handleCheckIn}
                  className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-800 py-3 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
                >
                  {status === "submitting" ? <RefreshCw size={15} className="animate-spin" /> : <UserCheck size={15} />}
                  {alreadyCheckedIn ? "Already Checked In" : "Confirm Check-In"}
                </button>
              </>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function AuthenticatedApp() {
  const { currentUser, logout, currentAdmin, adminLogout } = useAppData();
  const [view, setView] = useState("dashboard");

  const authed = Boolean(currentUser) || Boolean(currentAdmin);

  if (!authed) {
    return (
      <LoginPage
        onParticipantLogin={() => setView("dashboard")}
        onAdminLogin={() => setView("admin")}
      />
    );
  }

  const handleLogout = () => {
    logout();
    adminLogout();
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800">
      <TopNav view={view} setView={setView} onLogout={handleLogout} isAdmin={Boolean(currentAdmin)} />
      {view === "dashboard" && <Dashboard openAssessment={() => setView("profile")} />}
      {view === "profile" && <Profile />}
      {view === "surveys" && <ParticipantSurveysPage />}
      {view === "myGroup" && <ParticipantGroupPage />}
      {view === "myMeetings" && <ParticipantMeetingsPage />}
      {view === "myResources" && <ParticipantResourcesPage />}
      {view === "survey" && <SurveyAnalytics />}
      {view === "admin" && <Admin />}
    </div>
  );
}

function VerifyEmailPage({ token, email }) {
  const { verifyEmail } = useAppData();
  const [status, setStatus] = useState("verifying"); // verifying | done | error
  const [error, setError] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const result = await verifyEmail(email, token);
        if (!result.ok) {
          setError(result.error);
          setStatus("error");
        } else {
          setStatus("done");
        }
      } catch (err) {
        setError("Something went wrong. Please try again.");
        setStatus("error");
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const goToApp = () => {
    window.location.href = window.location.origin + window.location.pathname;
  };

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-50 p-6">
      <div className="w-full" style={{ maxWidth: 420 }}>
        <div className="mb-6 flex justify-center"><Logo size={36} /></div>
        <div className="rounded-2xl border border-slate-200 bg-white p-6 text-center">
          {status === "verifying" && (
            <div className="flex flex-col items-center py-4">
              <RefreshCw size={26} className="mb-3 animate-spin text-blue-500" />
              <p className="text-sm text-slate-500">Verifying your email…</p>
            </div>
          )}
          {status === "done" && (
            <div className="flex flex-col items-center">
              <CheckCircle2 size={30} className="mb-3 text-emerald-500" />
              <h2 className="text-lg font-bold text-slate-800">Email verified!</h2>
              <p className="mt-1 text-sm text-slate-500">You're signed in — head back to the app to continue.</p>
              <button onClick={goToApp} className="mt-5 flex items-center gap-2 rounded-lg bg-blue-800 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-900">
                Go to MercerAssess
              </button>
            </div>
          )}
          {status === "error" && (
            <div className="flex flex-col items-center">
              <XCircle size={30} className="mb-3 text-red-500" />
              <h2 className="text-lg font-bold text-slate-800">Couldn't verify</h2>
              <p className="mt-1 text-sm text-slate-500">{error}</p>
              <button onClick={goToApp} className="mt-5 flex items-center gap-2 rounded-lg bg-blue-800 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-900">
                Go to MercerAssess
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function ResetPasswordPage({ token, email, isAdmin }) {
  const { resetPassword, resetAdminPassword } = useAppData();
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [status, setStatus] = useState("form"); // form | submitting | done
  const [error, setError] = useState(null);

  const handleSubmit = async () => {
    if (password !== confirmPassword) { setError("Passwords don't match."); return; }
    setStatus("submitting");
    setError(null);
    try {
      const result = isAdmin ? await resetAdminPassword(email, token, password) : await resetPassword(email, token, password);
      if (!result.ok) {
        setError(result.error);
        setStatus("form");
      } else {
        setStatus("done");
      }
    } catch (err) {
      setError("Something went wrong. Please try again.");
      setStatus("form");
    }
  };

  const goToApp = () => {
    window.location.href = window.location.origin + window.location.pathname;
  };

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-50 p-6">
      <div className="w-full" style={{ maxWidth: 420 }}>
        <div className="mb-6 flex justify-center"><Logo size={36} /></div>
        <div className="rounded-2xl border border-slate-200 bg-white p-6">
          {status === "done" ? (
            <div className="flex flex-col items-center text-center">
              <CheckCircle2 size={30} className="mb-3 text-emerald-500" />
              <h2 className="text-lg font-bold text-slate-800">Password updated</h2>
              <p className="mt-1 text-sm text-slate-500">You're signed in — head back to the app to continue.</p>
              <button onClick={goToApp} className="mt-5 flex items-center gap-2 rounded-lg bg-blue-800 px-4 py-2.5 text-sm font-semibold text-white hover:bg-blue-900">
                Go to MercerAssess
              </button>
            </div>
          ) : (
            <>
              <h2 className="mb-1 text-lg font-bold text-slate-800">Set a new password</h2>
              <p className="mb-5 text-sm text-slate-400">for {email}</p>
              <div className="space-y-3">
                <div>
                  <label className="mb-1 block text-sm font-medium text-slate-700">New password</label>
                  <div className="relative">
                    <input
                      autoFocus
                      type={showPassword ? "text" : "password"}
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      className="w-full rounded-lg border border-slate-200 px-3 py-2.5 pr-10 text-sm outline-none focus:border-blue-400"
                    />
                    <button type="button" onClick={() => setShowPassword((s) => !s)} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                      {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                    </button>
                  </div>
                </div>
                <div>
                  <label className="mb-1 block text-sm font-medium text-slate-700">Confirm new password</label>
                  <input
                    type={showPassword ? "text" : "password"}
                    value={confirmPassword}
                    onChange={(e) => setConfirmPassword(e.target.value)}
                    className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-blue-400"
                  />
                </div>
                {error && <div className="rounded-lg bg-red-50 px-3 py-2.5 text-xs text-red-600">{error}</div>}
                <button
                  onClick={handleSubmit}
                  disabled={status === "submitting" || !password || !confirmPassword}
                  className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-800 py-3 text-sm font-semibold text-white hover:bg-blue-900 disabled:cursor-not-allowed disabled:opacity-40"
                >
                  {status === "submitting" ? <RefreshCw size={15} className="animate-spin" /> : <CheckCircle2 size={15} />}
                  Set New Password
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

export default function MercerAssess() {
  const params = typeof window !== "undefined" ? new URLSearchParams(window.location.search) : null;
  const surveyId = params ? params.get("survey") : null;
  const attendanceId = params ? params.get("attendance") : null;
  const resetToken = params ? params.get("resetToken") : null;
  const resetEmail = params ? params.get("resetEmail") : null;
  const adminResetToken = params ? params.get("adminResetToken") : null;
  const adminResetEmail = params ? params.get("adminResetEmail") : null;
  const verifyToken = params ? params.get("verifyToken") : null;
  const verifyEmailParam = params ? params.get("verifyEmail") : null;

  if (surveyId) {
    // Public survey-taking link (from a QR code or shared URL) — no login,
    // no participant data needed, just the survey form itself.
    return <TakeSurveyPage surveyId={surveyId} />;
  }

  if (attendanceId) {
    // Public attendance check-in link (from a QR code or shared URL). Wrapped
    // in AppDataProvider (unlike the survey link) because check-in now
    // requires signing in as yourself — this is what prevents one participant
    // checking in on behalf of another who isn't actually present.
    return (
      <AppDataProvider>
        <TakeAttendancePage sessionId={attendanceId} />
      </AppDataProvider>
    );
  }

  if (resetToken && resetEmail) {
    // Password reset link from an emailed reset request.
    return (
      <AppDataProvider>
        <ResetPasswordPage token={resetToken} email={resetEmail} />
      </AppDataProvider>
    );
  }

  if (adminResetToken && adminResetEmail) {
    return (
      <AppDataProvider>
        <ResetPasswordPage token={adminResetToken} email={adminResetEmail} isAdmin />
      </AppDataProvider>
    );
  }

  if (verifyToken && verifyEmailParam) {
    return (
      <AppDataProvider>
        <VerifyEmailPage token={verifyToken} email={verifyEmailParam} />
      </AppDataProvider>
    );
  }

  return (
    <AppDataProvider>
      <AuthenticatedApp />
    </AppDataProvider>
  );
}
