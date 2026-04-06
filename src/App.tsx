/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileText, 
  BarChart3, 
  PieChart as PieChartIcon, 
  TrendingUp, 
  Download, 
  CheckCircle2, 
  AlertCircle,
  Info,
  ChevronRight,
  Plus,
  Trash2,
  Table,
  FileDown,
  FileSpreadsheet,
  Printer,
  Loader2,
  HelpCircle,
  X
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  PieChart, 
  Pie, 
  Cell, 
  ScatterChart, 
  Scatter, 
  ZAxis,
  ReferenceLine,
  Legend
} from 'recharts';
import { motion, AnimatePresence } from 'motion/react';
import html2canvas from 'html2canvas';

import { cn } from '@/src/lib/utils';

// --- Types ---

interface CourseInfo {
  courseName: string;
  courseId: string;
  semester: string;
  courseNature: string;
  credits: number;
  classHours: number;
  examType: string;
  bookType: string;
  usualWeight: number;
  examWeight: number;
  objectiveRatios: number[]; // Ratios for Objective 1, 2, 3, 4
  usualAssignmentWeights: number[]; // Weights for Assignment 1, 2, 3, 4
  usualAssignmentObjectives: number[][]; // Objective indices for each assignment
  teacher: string;
  className: string;
  studentCount: number;
  objectiveDescriptions: string[];
}

interface ScoreRecord {
  id: string;
  name: string;
  className: string;
  examScores: Record<string, number>; // Question ID -> Score
  usualScores: number[]; // Scores for Assignment 1, 2, 3, 4
  examTotal: number;
  usualTotal: number;
  comprehensiveScore: number;
  objectiveScores: number[]; // Scores for Objective 1, 2, 3, 4
  objectiveAchievements: number[]; // Achievement values (0-1)
}

interface QuestionMapping {
  questionId: string;
  maxScore: number;
  objectiveIndex: number; // 0, 1, 2, 3
}

// --- Constants ---

const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'];
const GRADE_COLORS = {
  '优秀': '#10b981',
  '良好': '#3b82f6',
  '中等': '#f59e0b',
  '及格': '#8b5cf6',
  '不及格': '#ef4444',
};

// --- Main Component ---

export default function App() {
  // State
  const [courseInfo, setCourseInfo] = useState<CourseInfo>({
    courseName: '',
    courseId: '',
    semester: '',
    courseNature: '',
    credits: 0,
    classHours: 0,
    examType: '',
    bookType: '',
    usualWeight: 0,
    examWeight: 100,
    objectiveRatios: [],
    usualAssignmentWeights: [],
    usualAssignmentObjectives: [],
    teacher: '',
    className: '',
    studentCount: 0,
    objectiveDescriptions: []
  });

  const [questionMappings, setQuestionMappings] = useState<QuestionMapping[]>([]);
  const [students, setStudents] = useState<ScoreRecord[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [uploadedFiles, setUploadedFiles] = useState<{
    structure?: string;
    detail?: string;
    usual?: string;
  }>({});
  const [validationError, setValidationError] = useState<string[] | null>(null);
  const [showConfirm, setShowConfirm] = useState<{
    message: string;
    onConfirm: () => void;
  } | null>(null);
  const [showGuide, setShowGuide] = useState(false);
  const [guideStep, setGuideStep] = useState(0);

  const guideSteps = [
    {
      title: "欢迎使用分析系统",
      content: "这是一个专业的课程目标达成度分析工具。我们将引导您完成数据分析和报告生成的全过程。",
      targetId: "header-title",
    },
    {
      title: "第一步：填写课程信息",
      content: "在这里输入课程的基本信息。",
      targetId: "section-course-info",
    },
    {
      title: "第二步：设定课程目标",
      content: "定义您的课程目标描述。您可以根据需要添加或删除目标。",
      targetId: "section-objectives",
    },
    {
      title: "第三步：上传学生成绩",
      content: "下载模板并按格式填入学生成绩，然后上传 Excel 文件。系统会自动计算各项达成度。",
      targetId: "section-upload",
    },
    {
      title: "第四步：查看分析结果",
      content: "系统将生成详细的达成度分析表和可视化图表，帮助您直观了解教学效果。",
      targetId: "section-analysis",
    },
    {
      title: "最后：导出报告",
      content: "确认无误后，您可以一键导出精美的 Word 格式的分析报告。",
      targetId: "header-actions",
    }
  ];

  const nextGuideStep = () => {
    if (guideStep < guideSteps.length - 1) {
      setGuideStep(guideStep + 1);
      // Scroll to target
      const target = document.getElementById(guideSteps[guideStep + 1].targetId);
      if (target) {
        target.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    } else {
      setShowGuide(false);
      setGuideStep(0);
    }
  };

  const prevGuideStep = () => {
    if (guideStep > 0) {
      setGuideStep(guideStep - 1);
      const target = document.getElementById(guideSteps[guideStep - 1].targetId);
      if (target) {
        target.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }
  };

  const startGuide = () => {
    setGuideStep(0);
    setShowGuide(true);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };
  const [loadingMessage, setLoadingMessage] = useState('正在处理数据，请稍候...');
  const reportRef = useRef<HTMLDivElement>(null);

  const [showDataTable, setShowDataTable] = useState(false);
  const [evaluationMeasures, setEvaluationMeasures] = useState({
    evaluation: '',
    improvement: '',
    professorOpinion: '',
    auditOpinion: '',
    evaluationDate: '',
    auditDate: ''
  });

  // --- Constants & Aliases ---
  const ID_ALIASES = ['学号', '学生学号', 'ID', 'Student ID', '学 号', '学号(Student ID)'];
  const NAME_ALIASES = ['姓名', '学生姓名', 'Name', '姓 名', '姓名(Name)'];
  const CLASS_ALIASES = ['班级', '行政班', '专业班级', '班 级', '行政班级', '班级名称'];
  const USUAL_SCORE_ALIASES = ['平时总评', '总评', '平时成绩', '总分', '平时', '平时分', '平时成绩总分', '平时分合计', '综合平时', '平时成绩60%', '平时成绩(60%)'];
  const EXAM_SCORE_ALIASES = ['总分', '卷面成绩', '期末成绩', '考试成绩', '卷面', '期末', '卷面总分', '期末考试成绩', '期末总分'];
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [errorType, setErrorType] = useState<'upload' | 'export' | 'general'>('general');

  const parseScore = (val: any): number => {
    if (val === undefined || val === null || val === '') return 0;
    if (typeof val === 'number') return val;
    const s = String(val).trim();
    const num = parseFloat(s);
    if (!isNaN(num)) return num;
    
    const levels: Record<string, number> = {
      '优秀': 95, '优': 95,
      '良好': 85, '良': 85,
      '中等': 75, '中': 75,
      '及格': 65,
      '不及格': 50, '不合格': 50, '差': 50
    };
    for (const [level, score] of Object.entries(levels)) {
      if (s.includes(level)) return score;
    }
    return 0;
  };

  const getValByAliases = (row: any, aliases: string[]) => {
    const keys = Object.keys(row);
    for (const alias of aliases) {
      const normalizedAlias = alias.toLowerCase().replace(/\s/g, '');
      const foundKey = keys.find(k => {
        const normalizedK = k.toLowerCase().replace(/\s/g, '');
        return normalizedK === normalizedAlias || normalizedK.includes(normalizedAlias);
      });
      if (foundKey) return row[foundKey];
    }
    return undefined;
  };

  const robustParseExcel = (ws: XLSX.WorkSheet) => {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
    let headerRowIndex = -1;
    
    // Find the row containing "学号" or "姓名" or "题号"
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      if (rows[i] && rows[i].some(cell => {
        const s = String(cell || '').trim();
        return ID_ALIASES.includes(s) || NAME_ALIASES.includes(s) || s.includes('学号') || s.includes('姓名') || s.includes('题号');
      })) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) return [];

    const headers = rows[headerRowIndex].map(h => String(h || '').trim());
    const dataRows = rows.slice(headerRowIndex + 1);

    return dataRows
      .filter(row => row.length > 0 && (row.some(cell => cell !== null && cell !== undefined && cell !== '')))
      .map(row => {
        const obj: any = {};
        headers.forEach((h, idx) => {
          if (h) obj[h] = row[idx];
        });
        return obj;
      });
  };

  const addAssignment = () => {
    setCourseInfo(prev => {
      const newWeights = [...prev.usualAssignmentWeights, 0];
      // Recalculate the last weight to ensure sum is 100
      if (newWeights.length > 1) {
        const sumOfOthers = newWeights.slice(0, -1).reduce((sum, w) => sum + w, 0);
        newWeights[newWeights.length - 1] = Math.max(0, 100 - sumOfOthers);
      } else {
        newWeights[0] = 100;
      }
      
      return {
        ...prev,
        usualAssignmentWeights: newWeights,
        usualAssignmentObjectives: [...prev.usualAssignmentObjectives, []]
      };
    });
  };

  const removeAssignment = (index: number) => {
    if (courseInfo.usualAssignmentWeights.length <= 1) {
      setUploadError('至少需要保留一个平时作业项。');
      setErrorType('general');
      return;
    }
    setCourseInfo(prev => {
      const newWeights = prev.usualAssignmentWeights.filter((_, i) => i !== index);
      // Recalculate the last weight to ensure sum is 100
      if (newWeights.length > 1) {
        const sumOfOthers = newWeights.slice(0, -1).reduce((sum, w) => sum + w, 0);
        newWeights[newWeights.length - 1] = Math.max(0, 100 - sumOfOthers);
      } else if (newWeights.length === 1) {
        newWeights[0] = 100;
      }
      
      return {
        ...prev,
        usualAssignmentWeights: newWeights,
        usualAssignmentObjectives: prev.usualAssignmentObjectives.filter((_, i) => i !== index)
      };
    });
  };

  const formatClassNames = (classNames: string[]): string => {
    if (classNames.length === 0) return '未知班级';
    
    // Group by prefix (e.g., "道路")
    const groups: Record<string, number[]> = {};
    const nonStandard: string[] = [];
    
    classNames.forEach(name => {
      // Match something like "道路2201班", "道路2201", "道路 2201"
      // We look for a prefix followed by digits, optionally ending with '班'
      const match = name.match(/^(.+?)(\d+)(班)?$/);
      if (match) {
        const prefix = match[1].trim();
        const num = parseInt(match[2]);
        if (!groups[prefix]) groups[prefix] = [];
        if (!groups[prefix].includes(num)) groups[prefix].push(num);
      } else {
        // Only add to nonStandard if it looks like a real class name
        const keywords = ['班级', '行政班', '比例', '合计', '学号', '姓名', '成绩', '考试', '页码', '打印'];
        if (name.length >= 2 && !keywords.some(k => name.includes(k)) && !nonStandard.includes(name)) {
          nonStandard.push(name);
        }
      }
    });

    const result: string[] = [];
    
    Object.entries(groups).forEach(([prefix, nums]) => {
      nums.sort((a, b) => a - b);
      
      const sequences: { start: number, end: number }[] = [];
      if (nums.length > 0) {
        let currentSeq = { start: nums[0], end: nums[0] };
        
        for (let i = 1; i < nums.length; i++) {
          if (nums[i] === currentSeq.end + 1) {
            currentSeq.end = nums[i];
          } else {
            sequences.push(currentSeq);
            currentSeq = { start: nums[i], end: nums[i] };
          }
        }
        sequences.push(currentSeq);
      }
      
      sequences.forEach(seq => {
        if (seq.start === seq.end) {
          result.push(`${prefix}${seq.start}班`);
        } else {
          result.push(`${prefix}${seq.start}-${seq.end}班`);
        }
      });
    });
    
    // Combine with non-standard names
    const finalResult = [...result, ...nonStandard];
    return finalResult.join('、');
  };

  const getAssignmentAliases = (index: number) => {
    const num = index + 1;
    const chineseNums = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
    const chineseNum = chineseNums[index] || num;
    return [`作业${num}`, `第${num}次作业`, `作业${chineseNum}`, `第${chineseNum}次作业`, `Assignment ${num}`];
  };

  const handleExamStructureUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFiles(prev => ({ ...prev, structure: file.name }));
    setIsProcessing(true);
    setLoadingMessage('正在解析考试结构表...');
    setUploadError(null);
    setErrorType('upload');
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const dataArr = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(dataArr, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        
        // Get raw rows to check for horizontal layout
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        let newQuestionMappings: QuestionMapping[] = [];

        // Check for horizontal layout (like the image)
        let qRowIdx = -1, objRowIdx = -1, maxRowIdx = -1;
        for (let i = 0; i < Math.min(rows.length, 10); i++) {
          const firstCell = String(rows[i][0] || '').trim();
          if (firstCell.includes('试题') || firstCell.includes('题号')) qRowIdx = i;
          if (firstCell.includes('对应课程目标') || firstCell.includes('课程目标')) objRowIdx = i;
          if (firstCell.includes('满分值') || firstCell.includes('满分')) maxRowIdx = i;
        }

        if (qRowIdx !== -1 && objRowIdx !== -1 && maxRowIdx !== -1) {
          // Horizontal layout detected
          const qRow = rows[qRowIdx];
          const objRow = rows[objRowIdx];
          const maxRow = rows[maxRowIdx];
          
          for (let j = 1; j < qRow.length; j++) {
            const qId = String(qRow[j] || '').trim();
            if (!qId || qId === '总分' || qId === '合计') continue;
            
            const maxScore = parseScore(maxRow[j]);
            const objStr = String(objRow[j] || '');
            const objMatch = objStr.match(/目标\s*(\d+)/) || objStr.match(/(\d+)/);
            const objIdx = objMatch ? parseInt(objMatch[1]) - 1 : 0;
            
            newQuestionMappings.push({
              questionId: qId,
              maxScore,
              objectiveIndex: objIdx
            });
          }
        } else {
          // Fallback to vertical layout
          const data = robustParseExcel(ws);
          if (data.length > 0) {
            newQuestionMappings = data.map(row => {
              const qId = String(getValByAliases(row, ['题号', '题目', 'Question ID']) || '').trim();
              const maxScore = parseScore(getValByAliases(row, ['满分', '分值', 'Max Score']));
              const objStr = String(getValByAliases(row, ['对应课程目标', '课程目标', 'Objective']) || '');
              const objMatch = objStr.match(/目标\s*(\d+)/) || objStr.match(/(\d+)/);
              const objIdx = objMatch ? parseInt(objMatch[1]) - 1 : 0;

              return {
                questionId: qId,
                maxScore,
                objectiveIndex: objIdx
              };
            }).filter(m => m.questionId !== '');
          }
        }

        if (newQuestionMappings.length === 0) {
          setUploadError('无法识别考试结构表。请确保表格包含“题号”、“满分”和“对应课程目标”。');
          setIsProcessing(false);
          return;
        }

        setQuestionMappings(newQuestionMappings);

        // Update courseInfo ratios based on new mappings
        const maxObjIdx = Math.max(...newQuestionMappings.map(m => m.objectiveIndex));
        const count = Math.max(courseInfo.objectiveRatios.length, maxObjIdx + 1);
        const newRatios = new Array(count).fill(0);
        const totalMaxScore = newQuestionMappings.reduce((sum, m) => sum + m.maxScore, 0);
        
        if (totalMaxScore > 0) {
          newQuestionMappings.forEach(m => {
            newRatios[m.objectiveIndex] += (m.maxScore / totalMaxScore) * 100;
          });
        }

        setCourseInfo(prev => ({
          ...prev,
          objectiveRatios: newRatios.map(r => Math.round(r)),
          objectiveDescriptions: new Array(count).fill('').map((_, i) => prev.objectiveDescriptions[i] || `课程目标${i + 1}描述...`)
        }));

      } catch (err) {
        setUploadError('解析考试结构表时出错，请检查文件格式。');
        console.error(err);
      }
      setIsProcessing(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleExamDetailUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFiles(prev => ({ ...prev, detail: file.name }));
    setIsProcessing(true);
    setLoadingMessage('正在解析考试得分明细表...');
    setUploadError(null);
    setErrorType('upload');
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const dataArr = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(dataArr, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = robustParseExcel(ws);

        if (data.length === 0) {
          setUploadError('无法识别考试得分明细表。请确保表格包含“学号”和“姓名”列。');
          setIsProcessing(false);
          return;
        }

        // Find class names from the entire sheet (more robust)
        const classNamesSet = new Set<string>();
        const fullRows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        // 1. First, try to grab class names from the data rows (the "Class" column)
        // This is usually the most accurate source if it exists in the student list
        data.forEach(row => {
          const val = getValByAliases(row, CLASS_ALIASES);
          if (val) {
            const s = String(val).trim();
            // Filter out common non-class strings that might appear in the column
            const keywords = ['班级', '行政班', '比例', '合计', '学号', '姓名', '成绩', '考试', '页码', '打印', '分数', '权重', '平均', '最高', '最低'];
            if (s.length >= 2 && !keywords.some(k => s.includes(k))) {
              classNamesSet.add(s);
            }
          }
        });

        // 2. If no class names found in the column, fallback to searching headers/metadata
        if (classNamesSet.size === 0) {
          for (let rIdx = 0; rIdx < Math.min(fullRows.length, 20); rIdx++) {
            const row = fullRows[rIdx];
            if (!row) continue;
            for (let cIdx = 0; cIdx < row.length; cIdx++) {
              const cell = row[cIdx];
              const s = String(cell || '').trim();
              if (s.includes('班级') || s.includes('行政班')) {
                const match = s.match(/(?:班级|行政班)[:：]?\s*(.+)/);
                if (match && match[1].trim().length > 1) {
                  const names = match[1].trim().split(/[,，、\s]+/).filter(n => n.length > 1);
                  names.forEach(n => {
                    const keywords = ['比例', '合计', '学号', '姓名', '成绩', '考试', '平均', '最高', '最低'];
                    if (!keywords.some(k => n.includes(k))) {
                      classNamesSet.add(n);
                    }
                  });
                } else {
                  const nextCell = String(row[cIdx + 1] || '').trim();
                  const keywords = ['学号', '姓名', '比例', '成绩', '合计', '平均', '最高', '最低'];
                  if (nextCell.length > 1 && !keywords.some(k => nextCell.includes(k))) {
                    classNamesSet.add(nextCell);
                  }
                  if (rIdx < fullRows.length - 1) {
                    const belowCell = String(fullRows[rIdx + 1][cIdx] || '').trim();
                    if (belowCell.length > 1 && !keywords.some(k => belowCell.includes(k))) {
                      classNamesSet.add(belowCell);
                    }
                  }
                }
              }
            }
          }
        }

        const finalClassName = formatClassNames(Array.from(classNamesSet));
        
        setCourseInfo(prev => ({
          ...prev,
          className: finalClassName,
          studentCount: data.filter(r => getValByAliases(r, ID_ALIASES)).length
        }));

        // Identify question columns for score extraction
        // If questionMappings is already set, use those IDs. Otherwise, try to detect from keys.
        const keys = Object.keys(data[0] || {});
        const detectedQCols = keys.filter(k => k.includes('题') || k.includes('小题'));
        
        const qCols = questionMappings.length > 0 
          ? questionMappings.map(m => m.questionId)
          : detectedQCols;

        // Process student data
        const processedStudents: ScoreRecord[] = data
          .filter(row => getValByAliases(row, ID_ALIASES))
          .map(row => {
            const examScores: Record<string, number> = {};
            let calculatedTotal = 0;
            qCols.forEach(q => {
              const score = parseScore(row[q]);
              examScores[q] = score;
              calculatedTotal += score;
            });

            const directTotal = getValByAliases(row, EXAM_SCORE_ALIASES);
            const examTotal = (directTotal !== undefined && directTotal !== null && directTotal !== '') 
              ? parseScore(directTotal) 
              : calculatedTotal;

            return {
              id: String(getValByAliases(row, ID_ALIASES) || '').trim(),
              name: String(getValByAliases(row, NAME_ALIASES) || '').trim(),
              className: String(getValByAliases(row, CLASS_ALIASES) || ''),
              examScores,
              usualScores: [0, 0, 0, 0],
              examTotal,
              usualTotal: 0,
              comprehensiveScore: 0,
              objectiveScores: [0, 0, 0, 0],
              objectiveAchievements: [0, 0, 0, 0],
            };
          }).filter(s => s.id !== '');

        setStudents(prev => {
          if (prev.length === 0) return processedStudents;
          const merged = [...processedStudents];
          prev.forEach(oldStudent => {
            const newStudent = merged.find(s => s.id === oldStudent.id);
            if (newStudent) {
              newStudent.usualTotal = oldStudent.usualTotal;
              newStudent.usualScores = oldStudent.usualScores;
            }
          });
          return merged;
        });

      } catch (err) {
        setUploadError('解析考试得分明细表时出错，请检查文件格式。');
        console.error(err);
      }
      setIsProcessing(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleUsualUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFiles(prev => ({ ...prev, usual: file.name }));
    setIsProcessing(true);
    setLoadingMessage('正在解析平时成绩表...');
    setUploadError(null);
    setErrorType('upload');
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const dataArr = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(dataArr, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = robustParseExcel(ws);

        if (data.length === 0) {
          setUploadError('无法识别平时成绩表。请确保表格包含“学号”列，且表头在工作表的前 20 行内。');
          setIsProcessing(false);
          return;
        }

        // --- Dynamic Objective Extraction from Usual Scores ---
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        const headerRowIndex = rows.findIndex(r => r && r.some(cell => {
          const s = String(cell || '').trim();
          return ID_ALIASES.includes(s) || NAME_ALIASES.includes(s);
        }));

        if (headerRowIndex > 0) {
          const metadataRows = rows.slice(0, headerRowIndex);
          const headerRow = rows[headerRowIndex];
          
          // Try to find "课程目标" mappings for assignments
          let objRowIndex = -1;
          for (let i = metadataRows.length - 1; i >= 0; i--) {
            if (metadataRows[i].some(cell => String(cell || '').includes('目标'))) {
              objRowIndex = i;
              break;
            }
          }

          if (objRowIndex !== -1) {
            const objRow = metadataRows[objRowIndex];
            const newAssignmentObjectives: number[][] = [];
            const newAssignmentWeights: number[] = [];
            const detectedDescriptions: Record<number, string> = {};
            let maxObjIdx = -1;
            
            headerRow.forEach((key, colIdx) => {
              const keyStr = String(key || '').trim();
              if (keyStr.includes('作业') || keyStr.includes('实验') || keyStr.includes('平时')) {
                const objVal = String(objRow[colIdx] || '');
                // Handle multiple objectives like "1,2" or "目标1,2"
                const objPart = objVal.includes(':') || objVal.includes('：') 
                  ? objVal.split(/[:：]/)[0] 
                  : objVal;
                
                const matches = objPart.match(/\d+/g);
                if (matches) {
                  const indices = matches.map(m => parseInt(m) - 1);
                  newAssignmentObjectives.push(indices);
                  indices.forEach(idx => { if (idx > maxObjIdx) maxObjIdx = idx; });

                  // Try to extract description if present after colon
                  const descMatch = objVal.match(/[:：]\s*(.*)/);
                  if (descMatch && descMatch[1].trim().length > 2) {
                    indices.forEach(idx => {
                      detectedDescriptions[idx] = descMatch[1].trim();
                    });
                  }

                  // Try to find weight in the same column or header
                  const weightMatch = keyStr.match(/(\d+)%/) || objVal.match(/(\d+)%/);
                  newAssignmentWeights.push(weightMatch ? parseInt(weightMatch[1]) : 0);
                }
              }
            });

            // Ensure weights sum to 100 if we detected weights
            if (newAssignmentWeights.length > 1 && newAssignmentWeights.some(w => w > 0)) {
              const sumOfOthers = newAssignmentWeights.slice(0, -1).reduce((sum, w) => sum + w, 0);
              newAssignmentWeights[newAssignmentWeights.length - 1] = Math.max(0, 100 - sumOfOthers);
            } else if (newAssignmentWeights.length === 1) {
              newAssignmentWeights[0] = 100;
            }

            if (newAssignmentObjectives.length > 0) {
              setCourseInfo(prev => {
                const currentCount = prev.objectiveRatios.length;
                const newCount = Math.max(currentCount, maxObjIdx + 1);
                
                let updatedRatios = [...prev.objectiveRatios];
                let updatedDescs = [...prev.objectiveDescriptions];

                // Expand if needed
                if (newCount > currentCount) {
                  const extra = new Array(newCount - currentCount).fill(0);
                  updatedRatios = [...updatedRatios, ...extra];
                  updatedDescs = [...updatedDescs, ...new Array(newCount - currentCount).fill('').map((_, i) => `课程目标${currentCount + i + 1}描述...`)];
                }

                // Update descriptions if found
                Object.entries(detectedDescriptions).forEach(([idx, desc]) => {
                  updatedDescs[parseInt(idx)] = desc;
                });

                return {
                  ...prev,
                  objectiveRatios: updatedRatios,
                  objectiveDescriptions: updatedDescs,
                  usualAssignmentObjectives: newAssignmentObjectives,
                  usualAssignmentWeights: newAssignmentWeights.some(w => w > 0) ? newAssignmentWeights : prev.usualAssignmentWeights
                };
              });
            }
          }
        }

        const getAssignmentScoresFromRow = (row: any) => {
          return courseInfo.usualAssignmentWeights.map((_, i) => {
            const aliases = getAssignmentAliases(i);
            const score = getValByAliases(row, aliases);
            return score !== undefined && score !== null && score !== '' ? parseScore(score) : 0;
          });
        };

        const calculateUsualTotalFromScores = (scores: number[]) => {
          let total = 0;
          scores.forEach((score, i) => {
            total += score * (courseInfo.usualAssignmentWeights[i] / 100);
          });
          return Math.round(total * 100) / 100;
        };

        setStudents(prev => {
          if (prev.length === 0) {
            return data
              .filter(row => getValByAliases(row, ID_ALIASES))
              .map(row => {
                const usualScores = getAssignmentScoresFromRow(row);
                const directTotal = getValByAliases(row, USUAL_SCORE_ALIASES);
                const usualTotal = (directTotal !== undefined && directTotal !== null && directTotal !== '') 
                  ? parseScore(directTotal) 
                  : calculateUsualTotalFromScores(usualScores);

                return {
                  id: String(getValByAliases(row, ID_ALIASES) || '').trim(),
                  name: String(getValByAliases(row, NAME_ALIASES) || '').trim(),
                  className: String(getValByAliases(row, CLASS_ALIASES) || ''),
                  examScores: {},
                  usualScores,
                  examTotal: 0,
                  usualTotal,
                  comprehensiveScore: 0,
                  objectiveScores: [0, 0, 0, 0],
                  objectiveAchievements: [0, 0, 0, 0],
                };
              }).filter(s => s.id !== '');
          }

          const updated = [...prev];
          let matchCount = 0;
          data.forEach(row => {
            const studentId = String(getValByAliases(row, ID_ALIASES) || '').trim();
            if (!studentId) return;

            const student = updated.find(s => s.id === studentId);
            if (student) {
              const usualScores = getAssignmentScoresFromRow(row);
              const directTotal = getValByAliases(row, USUAL_SCORE_ALIASES);
              student.usualScores = usualScores;
              student.usualTotal = (directTotal !== undefined && directTotal !== null && directTotal !== '') 
                ? parseScore(directTotal) 
                : calculateUsualTotalFromScores(usualScores);
              matchCount++;
            }
          });
          
          if (matchCount === 0) {
            setUploadError('平时成绩表中的学号与考试成绩表不匹配，请检查学号是否一致。');
          }
          
          return updated;
        });
      } catch (err) {
        setUploadError('解析平时成绩表时出错，请检查文件格式。');
        console.error(err);
      }
      setIsProcessing(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const downloadUsualTemplate = () => {
    if (students.length === 0) {
      setUploadError('请先上传考试得分明细表，以便生成包含学生名单的平时成绩模板。');
      return;
    }

    const templateData = students.map(s => {
      const row: any = {
        '学号': s.id,
        '姓名': s.name,
        '行政班': s.className,
      };
      courseInfo.usualAssignmentWeights.forEach((w, i) => {
        row[`作业${i+1}(${w}%)`] = '';
      });
      row['平时总评'] = '';
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "平时成绩模板");
    
    // Auto-size columns
    const wscols = [
      { wch: 15 }, // 学号
      { wch: 10 }, // 姓名
      { wch: 20 }, // 行政班
      ...courseInfo.usualAssignmentWeights.map(() => ({ wch: 15 })),
      { wch: 12 }, // 平时总评
    ];
    ws['!cols'] = wscols;

    XLSX.writeFile(wb, `${courseInfo.courseName || '课程'}_平时成绩填写模板.xlsx`);
  };

  const validateBeforeExport = () => {
    const errors: string[] = [];

    // 1. 课程基本信息不完整
    const missingInfo = [];
    if (!courseInfo.courseName) missingInfo.push('课程名称');
    if (!courseInfo.courseId) missingInfo.push('课程编号');
    if (!courseInfo.semester) missingInfo.push('开课学期');
    if (!courseInfo.teacher) missingInfo.push('任课教师');
    if (!courseInfo.className) missingInfo.push('专业班级');
    
    if (missingInfo.length > 0) {
      errors.push(`课程基本信息不完整，请填写：${missingInfo.join('、')}。`);
    }

    // 课程目标信息空白
    if (courseInfo.objectiveDescriptions.length === 0 || courseInfo.objectiveDescriptions.some(d => !d.trim())) {
      errors.push('课程目标信息不能为空，请完善课程目标描述。');
    }

    // 2. 平时成绩组成权重模块
    if (courseInfo.usualWeight > 0) {
      if (courseInfo.usualAssignmentWeights.length === 0) {
        errors.push('平时成绩占比不为 0%，但尚未添加和设定平时成绩组成权重（如作业、实验等）。');
      } else {
        for (let i = 0; i < courseInfo.usualAssignmentWeights.length; i++) {
          const weight = courseInfo.usualAssignmentWeights[i];
          const objectives = courseInfo.usualAssignmentObjectives[i];
          
          if (weight === 0) {
            errors.push(`平时成绩组成权重错误：作业 ${i + 1} 的权重不能为 0。`);
          }
          
          if (!objectives || objectives.length === 0) {
            errors.push(`平时成绩组成权重错误：请为作业 ${i + 1} 设定对应的课程目标。`);
          }
        }
      }
    }

    // 3. 成绩数据输入部分缺失
    const missingFiles = [];
    if (questionMappings.length === 0) missingFiles.push('考试结构表');
    if (students.length === 0) missingFiles.push('考试得分明细表');
    
    const hasUsualData = students.some(s => s.usualTotal > 0);
    if (courseInfo.usualWeight > 0) {
      if (!hasUsualData) missingFiles.push('平时成绩表');
    } else {
      // usualWeight is 0
      if (hasUsualData) {
        errors.push("检测到已上传平时成绩数据，但未设定平时成绩占比（当前为0%），请检查权重设置或清除平时成绩数据。");
      }
    }

    if (missingFiles.length > 0) {
      errors.push(`成绩数据缺失：请上传${missingFiles.join('、')}。`);
    }

    // Check for duplicates
    if (uploadedFiles.structure && uploadedFiles.detail && uploadedFiles.structure === uploadedFiles.detail) {
      errors.push("成绩数据异常：考试结构表与考试得分明细表使用了同一个文件。");
    }
    if (uploadedFiles.detail && uploadedFiles.usual && uploadedFiles.detail === uploadedFiles.usual) {
      errors.push("成绩数据异常：考试得分明细表与平时成绩表使用了同一个文件。");
    }
    if (uploadedFiles.structure && uploadedFiles.usual && uploadedFiles.structure === uploadedFiles.usual) {
      errors.push("成绩数据异常：考试结构表与平时成绩表使用了同一个文件。");
    }

    return errors.length > 0 ? errors : null;
  };

  // --- Calculations ---

  const calculatedObjectiveRatios = useMemo(() => {
    const objCount = courseInfo.objectiveRatios.length;
    const ratios = new Array(objCount).fill(0);
    
    // Exam contribution
    const totalExamMax = questionMappings.reduce((sum, m) => sum + m.maxScore, 0);
    if (totalExamMax > 0) {
      questionMappings.forEach(m => {
        if (m.objectiveIndex < objCount) {
          ratios[m.objectiveIndex] += (m.maxScore / totalExamMax) * courseInfo.examWeight;
        }
      });
    }

    // Usual contribution
    courseInfo.usualAssignmentWeights.forEach((weight, i) => {
      const mappedObjs = courseInfo.usualAssignmentObjectives[i];
      if (mappedObjs && mappedObjs.length > 0) {
        const weightPerObj = (weight * courseInfo.usualWeight / 100) / mappedObjs.length;
        mappedObjs.forEach(objIdx => {
          if (objIdx < objCount) {
            ratios[objIdx] += weightPerObj;
          }
        });
      }
    });

    return ratios.map(r => Math.round(r * 10) / 10);
  }, [courseInfo.examWeight, courseInfo.usualWeight, courseInfo.usualAssignmentWeights, courseInfo.usualAssignmentObjectives, questionMappings, courseInfo.objectiveRatios.length]);

  // Sync objective ratios
  React.useEffect(() => {
    const isDifferent = calculatedObjectiveRatios.some((r, i) => r !== courseInfo.objectiveRatios[i]);
    if (isDifferent) {
      setCourseInfo(prev => ({
        ...prev,
        objectiveRatios: calculatedObjectiveRatios
      }));
    }
  }, [calculatedObjectiveRatios, courseInfo.objectiveRatios]);

  const finalData = useMemo(() => {
    if (students.length === 0) return [];

    // Calculate max scores per objective from Exam
    const objCount = courseInfo.objectiveRatios.length;
    const objMaxExamScores = new Array(objCount).fill(0);
    questionMappings.forEach(m => {
      if (m.objectiveIndex < objCount) {
        objMaxExamScores[m.objectiveIndex] += m.maxScore;
      }
    });

    return students.map(s => {
      const comp = (s.usualTotal * courseInfo.usualWeight / 100) + (s.examTotal * courseInfo.examWeight / 100);
      
      const objDetails = courseInfo.objectiveRatios.map((ratio, i) => {
        // Exam part for this objective
        let examScore = 0;
        let examMax = 0;
        questionMappings.forEach(m => {
          if (m.objectiveIndex === i) {
            examScore += s.examScores[m.questionId] || 0;
            examMax += m.maxScore;
          }
        });
        const examAch = examMax > 0 ? examScore / examMax : null;

        // Usual part for this objective
        let usualWeightedScoreSum = 0;
        let usualWeightSum = 0;
        courseInfo.usualAssignmentObjectives.forEach((objs, assignIdx) => {
          if (objs.includes(i)) {
            const weight = courseInfo.usualAssignmentWeights[assignIdx];
            const score = s.usualScores?.[assignIdx] || 0;
            usualWeightedScoreSum += score * weight;
            usualWeightSum += weight;
          }
        });
        const usualRawScore = usualWeightSum > 0 ? (usualWeightedScoreSum / usualWeightSum) : 0;
        const usualAch = usualWeightSum > 0 ? usualRawScore / 100 : null;

        // Combine
        let combinedAch = 0;
        if (examAch !== null && usualAch !== null) {
          combinedAch = (usualAch * courseInfo.usualWeight + examAch * courseInfo.examWeight) / 100;
        } else if (examAch !== null) {
          combinedAch = examAch;
        } else if (usualAch !== null) {
          combinedAch = usualAch;
        }

        return {
          achievement: combinedAch,
          examScore: examAch !== null ? examAch * ratio : 0,
          usualScore: usualRawScore, // 0-100 scale as requested
          totalScore: combinedAch * ratio // Contribution to total score (out of ratio)
        };
      });

      return {
        ...s,
        comprehensiveScore: Math.round(comp * 100) / 100,
        objectiveScores: objDetails,
        objectiveAchievements: objDetails.map(d => d.achievement),
        totalExamScore: s.examTotal,
        finalGrade: Math.round(comp * 100) / 100
      };
    });
  }, [students, courseInfo, questionMappings]);

  const stats = useMemo(() => {
    if (finalData.length === 0) return null;
    const scores = finalData.map(s => s.comprehensiveScore).sort((a, b) => a - b);
    const sum = scores.reduce((a, b) => a + b, 0);
    
    const grades = {
      '优秀': finalData.filter(s => s.comprehensiveScore >= 85).length,
      '良好': finalData.filter(s => s.comprehensiveScore >= 70 && s.comprehensiveScore < 85).length,
      '中等': finalData.filter(s => s.comprehensiveScore >= 60 && s.comprehensiveScore < 70).length,
      '不及格': finalData.filter(s => s.comprehensiveScore < 60).length,
    };

    return {
      max: Math.max(...scores),
      min: Math.min(...scores),
      avg: Math.round((sum / scores.length) * 100) / 100,
      median: scores[Math.floor(scores.length / 2)],
      passCount: finalData.filter(s => s.comprehensiveScore >= 60).length,
      failCount: grades['不及格'],
      grades
    };
  }, [finalData]);

  const objectiveStats = useMemo(() => {
    if (finalData.length === 0) return [];
    return courseInfo.objectiveRatios.map((_, i) => {
      const achievements = finalData.map(s => s.objectiveAchievements[i]);
      const avg = achievements.reduce((a, b) => a + b, 0) / achievements.length;
      const lowCount = achievements.filter(v => v < 0.65).length;
      const highCount = achievements.filter(v => v >= 0.8).length;
      
      let conclusion = '未达成';
      if (avg >= 0.85) conclusion = '优秀达成';
      else if (avg >= 0.75) conclusion = '良好达成';
      else if (avg >= 0.65) conclusion = '基本达成';

      return {
        index: i + 1,
        avg: Math.round(avg * 100) / 100,
        lowCount,
        highCount,
        conclusion,
        data: finalData.map((s, idx) => ({ x: idx + 1, y: s.objectiveAchievements[i], name: s.name }))
      };
    });
  }, [finalData, courseInfo.objectiveRatios]);

  // Auto-generate evaluation text and improvement measures when stats are available
  React.useEffect(() => {
    if (objectiveStats.length > 0) {
      const overallAvg = Math.round((objectiveStats.reduce((sum, obj) => sum + obj.avg, 0) / objectiveStats.length) * 100) / 100;
      
      let text = `本次课程目标达成评价值平均为 ${overallAvg}。`;
      
      const achievedCount = objectiveStats.filter(obj => obj.avg >= 0.65).length;
      if (achievedCount === objectiveStats.length) {
        text += `各课程目标均达到了预期教学效果（达成度均 ≥ 0.65）。`;
      } else {
        text += `共有 ${achievedCount} 个课程目标达到预期教学效果。`;
      }

      text += `从各课程目标达成度情况来看：`;
      objectiveStats.forEach((obj, idx) => {
        text += `课程目标 ${obj.index} 的达成度平均值为 ${obj.avg}（${obj.conclusion}）${idx === objectiveStats.length - 1 ? '。' : '；'}`;
      });

      const sortedStats = [...objectiveStats].sort((a, b) => a.avg - b.avg);
      const lowest = sortedStats[0];
      const highest = sortedStats[sortedStats.length - 1];

      text += `其中，课程目标 ${highest.index} 达成情况最好。`;
      if (lowest.avg < 0.7) {
        text += `课程目标 ${lowest.index} 的达成度相对较低（${lowest.avg}），在今后的教学中需进一步加强对该目标对应知识点的讲解与训练。`;
      }

      // Generate dynamic improvement measures
      const improvementMeasures: string[] = [];
      objectiveStats.forEach(obj => {
        const { index, avg, conclusion } = obj;
        const desc = courseInfo.objectiveDescriptions[index - 1] || '';
        
        const templates = {
          '优秀达成': [
            `针对课程目标${index}（${desc.substring(0, 15)}...），学生掌握情况极佳。后续将引入更具挑战性的案例分析，进一步拓宽学生的专业视野。`,
            `课程目标${index}达成度处于高位。计划在下一轮教学中增加前沿技术讲座，保持学生对该领域知识的探索热情。`,
            `鉴于课程目标${index}的优秀表现，将总结现有教学经验，并在其他相关章节推广启发式教学法。`,
            `课程目标${index}达成效果显著，学生对核心概念掌握扎实。未来将探索跨学科协作项目，提升综合应用能力。`
          ],
          '良好达成': [
            `课程目标${index}达成情况良好。未来将加强课堂互动，通过更多的小组讨论深化学生对复杂概念的理解。`,
            `针对课程目标${index}，计划优化作业设计，增加实践性环节，以巩固学生对理论知识的应用能力。`,
            `课程目标${index}表现稳健。将进一步完善教学大纲，确保知识点的衔接更加紧密。`,
            `该目标达成度较为理想。计划在后续课程中引入更多工程实际案例，强化理论联系实际的教学导向。`
          ],
          '基本达成': [
            `课程目标${index}仅为基本达成，反映出部分学生在“${desc.substring(0, 10)}”方面存在薄弱环节。需增加课后辅导频率。`,
            `针对课程目标${index}达成度偏低的问题，计划调整教学重点，增加基础知识的复习课时，并强化过程化考核。`,
            `课程目标${index}的达成度有待提高。将重新设计实验环节，提高学生的动手能力和解决实际问题的能力。`,
            `分析发现学生在目标${index}相关的复杂计算上存在短板。未来将通过增加课堂练习和专项测试来精准突破。`
          ],
          '未达成': [
            `课程目标${index}未达成，形势严峻。必须重新审视该部分的教学设计，分析学生普遍困惑的原因，并进行专题补课。`,
            `针对课程目标${index}未达标的情况，将组织任课教师集体研讨，改进教学手段，并对不及格学生进行重点帮扶。`,
            `课程目标${index}达成度极低。计划在下一学期大幅增加该部分的学时分配，并引入更多直观的教学辅助工具。`,
            `目标${index}的达成度未达预期。初步分析是由于前置知识衔接不畅，计划在开课初期增加前置知识的摸底与补强。`
          ]
        };
        
        const category = conclusion.includes('优秀') ? '优秀达成' : 
                         conclusion.includes('良好') ? '良好达成' : 
                         conclusion.includes('基本') ? '基本达成' : '未达成';
        
        const options = templates[category] || templates['基本达成'];
        const selected = options[Math.floor(Math.random() * options.length)];
        improvementMeasures.push(`${index}. ${selected}`);
      });
      const improvementText = improvementMeasures.join('\n');

      // Only auto-update if it's currently empty or looks like a previous auto-generated text
      setEvaluationMeasures(prev => {
        const newMeasures = { ...prev };
        if (prev.evaluation === '' || prev.evaluation.includes('达成评价值平均为')) {
          newMeasures.evaluation = text;
        }
        if (prev.improvement === '' || prev.improvement.includes('针对课程目标')) {
          newMeasures.improvement = improvementText;
        }
        return newMeasures;
      });
    }
  }, [objectiveStats]);

  // --- Actions ---

  const exportToExcel = () => {
    if (!finalData.length) return;

    const wb = XLSX.utils.book_new();
    const ws_data: any[][] = [];

    // Title
    ws_data.push(["重庆交通大学学生考核成绩统计表"]);
    ws_data.push([`（${courseInfo.semester}）`]);
    ws_data.push([]); // Empty row

    // Header Info
    ws_data.push([
      "学    院", "土木工程学院", "", 
      "课程名称", courseInfo.courseName, "", 
      "学    时", courseInfo.classHours, "", 
      "期末比例", `${courseInfo.examWeight}%`
    ]);
    ws_data.push([
      "专业班级", courseInfo.className, "", 
      "课程编号", courseInfo.courseId, "", 
      "学    分", courseInfo.credits, "", 
      "平时比例", `${courseInfo.usualWeight}%`
    ]);
    ws_data.push([]); // Empty row

    // Table Headers
    const headerRow1 = ["课程目标", "", ""];
    const headerRow2 = ["考核方式", "", ""];
    const headerRow3 = ["满分", "", ""];
    const headerRow4 = ["学号", "姓名", "班级"];

    courseInfo.objectiveRatios.forEach((ratio, i) => {
      headerRow1.push(`课程目标${i + 1}`, "", "");
      headerRow2.push("期末考试", "平时考核", "评价值");
      headerRow3.push(ratio.toString(), "100", "100");
      headerRow4.push(`T${String.fromCharCode(65 + i)}1`, `T${String.fromCharCode(65 + i)}2`, `T${String.fromCharCode(65 + i)}A`);
    });

    headerRow1.push("期末卷面成绩", "课程目标评价值");
    headerRow2.push("", "");
    headerRow3.push("100", "100");
    headerRow4.push("TTT", "TTA");

    ws_data.push(headerRow1);
    ws_data.push(headerRow2);
    ws_data.push(headerRow3);
    ws_data.push(headerRow4);

    // Student Data
    finalData.forEach(student => {
      const row = [student.id, student.name, student.className || courseInfo.className];
      student.objectiveScores.forEach(obj => {
        // examScore is out of ratio, usualScore is out of 100, evaluation value is out of 100 (integer)
        row.push(
          obj.examScore.toFixed(1), 
          obj.usualScore.toFixed(1), 
          Math.round(obj.achievement * 100).toString()
        );
      });
      row.push(
        student.totalExamScore.toFixed(1), 
        Math.round(student.finalGrade).toString()
      );
      ws_data.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    // Merges
    const merges = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: headerRow1.length - 1 } }, // Title
      { s: { r: 1, c: 0 }, e: { r: 1, c: headerRow1.length - 1 } }, // Subtitle
      // Header Info Merges
      { s: { r: 3, c: 1 }, e: { r: 3, c: 2 } },
      { s: { r: 3, c: 4 }, e: { r: 3, c: 5 } },
      { s: { r: 3, c: 7 }, e: { r: 3, c: 8 } },
      { s: { r: 4, c: 1 }, e: { r: 4, c: 2 } },
      { s: { r: 4, c: 4 }, e: { r: 4, c: 5 } },
      { s: { r: 4, c: 7 }, e: { r: 4, c: 8 } },
      // Table Header Merges (Left side)
      { s: { r: 6, c: 0 }, e: { r: 6, c: 2 } }, // 课程目标
      { s: { r: 7, c: 0 }, e: { r: 7, c: 2 } }, // 考核方式
      { s: { r: 8, c: 0 }, e: { r: 8, c: 2 } }, // 满分
    ];

    // Merge Objective Headers
    let currentCol = 3;
    courseInfo.objectiveRatios.forEach(() => {
      merges.push({ s: { r: 6, c: currentCol }, e: { r: 6, c: currentCol + 2 } });
      currentCol += 3;
    });

    // Merge Final Columns
    merges.push({ s: { r: 6, c: currentCol }, e: { r: 7, c: currentCol } }); // 期末卷面成绩
    merges.push({ s: { r: 6, c: currentCol + 1 }, e: { r: 7, c: currentCol + 1 } }); // 课程目标评价值

    ws['!merges'] = merges;

    XLSX.utils.book_append_sheet(wb, ws, "考核成绩统计表");
    XLSX.writeFile(wb, `考核成绩统计表_${courseInfo.courseName}_${courseInfo.className}.xlsx`);
  };

  const downloadChart = async (id: string, fileName: string) => {
    const element = document.getElementById(id);
    if (!element) return;
    
    try {
      // Find the SVG element within the container
      const svgElement = element.querySelector('svg');
      if (!svgElement) {
        console.error('No SVG found in chart container');
        return;
      }

      // Clone the SVG to modify it without affecting the UI
      const clonedSvg = svgElement.cloneNode(true) as SVGElement;
      
      // Ensure the SVG has explicit dimensions for sharp to render correctly
      const width = element.offsetWidth;
      const height = element.offsetHeight;
      clonedSvg.setAttribute('width', width.toString());
      clonedSvg.setAttribute('height', height.toString());
      
      // Add a white background if needed (Recharts SVGs are usually transparent)
      // We can wrap the content in a <rect> or just let sharp handle it if possible.
      // Sharp's default background is transparent. Let's add a white background rect.
      const backgroundRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      backgroundRect.setAttribute('width', '100%');
      backgroundRect.setAttribute('height', '100%');
      backgroundRect.setAttribute('fill', '#ffffff');
      clonedSvg.insertBefore(backgroundRect, clonedSvg.firstChild);

      const svgData = new XMLSerializer().serializeToString(clonedSvg);

      const response = await fetch('/api/render-chart', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          svg: svgData,
          width: width,
          height: height
        }),
      });

      if (!response.ok) {
        throw new Error('Failed to render chart on server');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${fileName}.png`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error downloading chart:', err);
    }
  };


  const exportWord = async () => {
    const error = validateBeforeExport();
    if (error) {
      setValidationError(error);
      return;
    }

    if (!reportRef.current) return;
    setIsProcessing(true);
    setIsExporting(true);
    setUploadError(null);
    setErrorType('export');
    setLoadingMessage('正在准备 Word 报告内容...');
    
    // Give charts time to settle
    await new Promise(resolve => setTimeout(resolve, 3000));
    window.scrollTo(0, 0);
    
    try {
      // We'll use the HTML-to-Word approach (MIME type application/msword)
      // This allows Word to open the HTML content and render it.
      
      const header = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' 
              xmlns:w='urn:schemas-microsoft-com:office:word' 
              xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <meta charset='utf-8'>
          <title>达成度分析报告</title>
          <!--[if gte mso 9]>
          <xml>
            <w:WordDocument>
              <w:View>Print</w:View>
              <w:Zoom>100</w:Zoom>
              <w:DoNotOptimizeForBrowser/>
            </w:WordDocument>
          </xml>
          <![endif]-->
          <style>
            @page {
              size: A4;
              margin: 1in;
            }
            body { 
              font-family: 'Times New Roman', '宋体', 'SimSun', serif; 
              line-height: 150%; 
              font-size: 12pt; 
              mso-pagination: widow-orphan;
            }
            table { 
              border-collapse: collapse; 
              width: 100%; 
              table-layout: auto; 
              margin-bottom: 20px; 
              border: 1px solid #000; 
              mso-table-lspace: 0pt; 
              mso-table-rspace: 0pt;
              mso-border-alt: solid windowtext .5pt;
            }
            th, td { 
              border: 1px solid #000; 
              padding: 4px; 
              text-align: center; 
              font-family: 'Times New Roman', '宋体', 'SimSun', serif; 
              font-size: 9pt; 
              line-height: 150%; 
              mso-border-alt: solid windowtext .5pt;
            }
            th { 
              font-weight: bold; 
              background-color: #f2f2f2; 
            }
            h1, h2, h3, h4, h5, h6 {
              font-family: 'Times New Roman', '黑体', 'Microsoft YaHei', sans-serif; 
              font-weight: bold; 
              line-height: 150%;
              margin-top: 12pt;
              margin-bottom: 6pt;
            }
            h1 { text-align: center; font-size: 18pt; }
            h2 { font-size: 16pt; }
            h3 { font-size: 14pt; }
            h4 { font-size: 12pt; }
            p, li, div { 
              line-height: 150%;
            }
            p { 
              margin-bottom: 6pt; 
              text-align: justify; 
              text-justify: inter-ideograph; 
              text-indent: 2em; 
              mso-char-indent-count: 2.0;
              font-size: 12pt;
            }
            ul, ol {
              margin-top: 0;
              margin-bottom: 6pt;
              line-height: 150%;
            }
            .mark { 
              background-color: #ffff00; 
              mso-highlight: yellow;
              font-weight: bold; 
            }
            .bg-slate-50 { background-color: #f2f2f2; }
            .font-bold { font-weight: bold; }
            .text-center { text-align: center; }
            .text-left { text-align: left; text-indent: 0; mso-char-indent-count: 0; }
            .break-before-page { page-break-before: always; }
            .chart-caption { 
              text-align: center; 
              font-size: 10.5pt; 
              margin-top: 4pt; 
              margin-bottom: 12pt; 
              font-family: 'Times New Roman', '宋体', 'SimSun', serif; 
              text-indent: 0; 
              mso-char-indent-count: 0;
              line-height: 150%;
            }
            img { 
              display: block; 
              margin: 10pt auto; 
              width: 14cm; 
              height: auto; 
            }
            .text-red-600 { color: #dc2626; font-weight: bold; }
            .text-emerald-600 { color: #059669; font-weight: bold; }
            .text-blue-600 { color: #2563eb; font-weight: bold; }
            .text-blue-700 { color: #1d4ed8; font-weight: bold; }
            /* Ensure no indent for elements inside table or specific containers */
            td, th { line-height: 150%; font-size: 9pt; }
            td p, th p, td div, th div { 
              text-indent: 0; 
              margin: 0; 
              mso-char-indent-count: 0; 
              line-height: 150%; 
              font-size: 9pt;
            }
          </style>
        </head>
        <body>
      `;
      
      const footer = "</body></html>";
      
      // Clone the report content to manipulate it for Word
      const reportClone = reportRef.current.cloneNode(true) as HTMLElement;
      
      // Reset main container styles for Word
      reportClone.style.width = '100%';
      reportClone.style.maxWidth = 'none';
      reportClone.style.padding = '0';
      reportClone.style.margin = '0';
      reportClone.className = ''; // Remove all Tailwind classes from the root clone
      
      // Ensure all tables have width="100%" attribute for Word
      const tables = reportClone.querySelectorAll('table');
      tables.forEach(table => {
        table.setAttribute('width', '100%');
        table.style.width = '100%';
        // Remove any fixed widths that might cause overflow
        table.style.minWidth = 'auto';
        table.style.maxWidth = '100%';
      });

      // Ensure all existing images have width="14cm" attribute
      const images = reportClone.querySelectorAll('img');
      images.forEach(img => {
        img.setAttribute('width', '529'); // Approx 14cm in pixels (96dpi)
        img.style.width = '14cm';
        img.style.height = 'auto';
        img.style.maxWidth = '100%';
        img.style.display = 'block';
        img.style.margin = '10pt auto';
      });

      // Remove fixed widths from common containers
      const containers = reportClone.querySelectorAll('div, section');
      containers.forEach(el => {
        const style = (el as HTMLElement).style;
        if (style.width && style.width !== '100%') style.width = '100%';
        style.maxWidth = '100%';
        // Remove horizontal padding that might push content out
        style.paddingLeft = '0';
        style.paddingRight = '0';
      });
      
      // Remove elements that shouldn't be in Word
      const ignores = reportClone.querySelectorAll('[data-html2canvas-ignore]');
      ignores.forEach(el => el.remove());
      
      // Convert charts to images because Word can't render Recharts
      const chartContainersInClone = Array.from(reportClone.querySelectorAll('[id^="chart-"]')) as HTMLElement[];
      console.log(`Found ${chartContainersInClone.length} chart containers in clone`);
      
      for (let i = 0; i < chartContainersInClone.length; i++) {
        const cloneContainer = chartContainersInClone[i];
        const id = cloneContainer.id;
        const originalContainer = document.getElementById(id);
        
        if (originalContainer) {
          setLoadingMessage(`正在转换图表 ${i + 1}/${chartContainersInClone.length}...`);
          try {
            const svgElement = originalContainer.querySelector('svg');
            if (svgElement) {
              // Clone the SVG to avoid modifying the original
              const clonedSvg = svgElement.cloneNode(true) as SVGSVGElement;
              
              // Get original dimensions from the container to ensure no cropping
              const rect = originalContainer.getBoundingClientRect();
              const originalWidth = rect.width || originalContainer.scrollWidth || 800;
              const originalHeight = rect.height || originalContainer.scrollHeight || 400;
              
              // Ensure viewBox is set correctly to capture everything
              // We use a slightly larger viewBox to be safe
              clonedSvg.setAttribute('viewBox', `0 0 ${originalWidth} ${originalHeight}`);
              
              // Set explicit width and height for the export (high res)
              const exportWidth = 2000;
              // Use a more consistent aspect ratio for Word (around 2:1 or 16:9)
              // If the original aspect ratio is too tall, we cap it.
              let aspectRatio = originalHeight / originalWidth;
              if (id.startsWith('chart-objective-')) {
                // For objective charts, we want them to be wide and clear
                aspectRatio = 0.45; // Roughly 2.2:1
              } else if (id === 'chart-scatter') {
                aspectRatio = 0.6; // Roughly 1.6:1
              } else {
                aspectRatio = Math.min(aspectRatio, 0.6); // Cap at 1.6:1
              }
              const exportHeight = exportWidth * aspectRatio;
              clonedSvg.setAttribute('width', exportWidth.toString());
              clonedSvg.setAttribute('height', exportHeight.toString());
              
              // Remove any clip-paths that might be causing cropping
              const clipPaths = clonedSvg.querySelectorAll('clipPath');
              clipPaths.forEach(cp => cp.remove());
              
              // Also remove clip-path attributes from all elements
              const allElements = clonedSvg.querySelectorAll('*');
              allElements.forEach(el => {
                if (el.hasAttribute('clip-path')) {
                  el.removeAttribute('clip-path');
                }
              });

              clonedSvg.style.overflow = 'visible';
              clonedSvg.style.backgroundColor = 'white'; // Ensure background is white
              
              // Add a white background rectangle
              const bgRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
              bgRect.setAttribute('x', '-10%'); // Extra coverage
              bgRect.setAttribute('y', '-10%');
              bgRect.setAttribute('width', '120%');
              bgRect.setAttribute('height', '120%');
              bgRect.setAttribute('fill', 'white');
              clonedSvg.insertBefore(bgRect, clonedSvg.firstChild);

              const svgData = new XMLSerializer().serializeToString(clonedSvg);
              
              const response = await fetch('/api/render-chart', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ svgData, width: exportWidth, height: exportHeight }),
              });

              if (response.ok) {
                const blob = await response.blob();
                const reader = new FileReader();
                const base64Promise = new Promise<string>((resolve) => {
                  reader.onloadend = () => resolve(reader.result as string);
                  reader.readAsDataURL(blob);
                });
                const imgData = await base64Promise;
                
                const img = document.createElement('img');
                img.src = imgData;
                img.setAttribute('width', '529'); // Approx 14cm in pixels (96dpi)
                img.style.width = '14cm';
                img.style.height = 'auto';
                img.style.display = 'block';
                img.style.margin = '10pt auto';
                
                // Replace the chart container content with the image
                cloneContainer.innerHTML = '';
                cloneContainer.style.height = 'auto';
                cloneContainer.style.width = '100%';
                cloneContainer.style.padding = '0';
                cloneContainer.style.margin = '0';
                cloneContainer.style.background = 'none';
                cloneContainer.style.border = 'none';
                // Add a caption
                let captionText = '';
                if (id === 'chart-histogram') captionText = '综合成绩分布直方图';
                else if (id === 'chart-pie') captionText = '各分数段占比饼图';
                else if (id === 'chart-scatter') captionText = '平时成绩 vs 考试成绩散点图';
                else if (id.startsWith('chart-objective-')) {
                  const objIndex = parseInt(id.replace('chart-objective-', ''));
                  captionText = `图${objIndex + 1} 课程目标${objIndex + 1}学生达成情况分布图`;
                }

                if (captionText) {
                  const caption = document.createElement('div');
                  caption.className = 'chart-caption';
                  caption.style.width = '100%';
                  caption.style.textAlign = 'center';
                  caption.textContent = captionText;
                  cloneContainer.appendChild(caption);
                }
                
                console.log(`Successfully converted chart ${id} to image via backend`);
              } else {
                console.warn(`Backend failed to render chart ${id}, falling back to html2canvas`);
                throw new Error('Backend rendering failed');
              }
            } else {
              throw new Error('No SVG found');
            }
          } catch (err) {
            console.warn(`Error using backend for chart ${id}, falling back to html2canvas:`, err);
            try {
              const canvas = await html2canvas(originalContainer, { 
                scale: 3,
                useCORS: true,
                allowTaint: true,
                logging: false,
                backgroundColor: '#ffffff',
                imageTimeout: 30000,
                scrollX: 0,
                scrollY: -window.scrollY,
                width: originalContainer.scrollWidth,
                height: originalContainer.scrollHeight
              });
              
              if (canvas.width > 0 && canvas.height > 0) {
                const imgData = canvas.toDataURL('image/png');
                const img = document.createElement('img');
                img.src = imgData;
                img.setAttribute('width', '529');
                img.style.width = '14cm';
                img.style.height = 'auto';
                img.style.display = 'block';
                img.style.margin = '10pt auto';
                
                cloneContainer.innerHTML = '';
                cloneContainer.style.height = 'auto';
                cloneContainer.style.width = '100%';
                cloneContainer.style.padding = '0';
                cloneContainer.style.margin = '0';
                cloneContainer.style.background = 'none';
                cloneContainer.style.border = 'none';
                cloneContainer.appendChild(img);

                let captionText = '';
                if (id === 'chart-histogram') captionText = '综合成绩分布直方图';
                else if (id === 'chart-pie') captionText = '各分数段占比饼图';
                else if (id === 'chart-scatter') captionText = '平时成绩 vs 考试成绩散点图';
                else if (id.startsWith('chart-objective-')) {
                  const objIndex = parseInt(id.replace('chart-objective-', ''));
                  captionText = `图${objIndex + 1} 课程目标${objIndex + 1}学生达成情况分布图`;
                }

                if (captionText) {
                  const caption = document.createElement('div');
                  caption.className = 'chart-caption';
                  caption.textContent = captionText;
                  cloneContainer.appendChild(caption);
                }
              }
            } catch (fallbackErr) {
              console.error(`Fallback failed for chart ${id}:`, fallbackErr);
            }
          }
        }
        // Small delay to prevent UI freezing
        await new Promise(resolve => setTimeout(resolve, 50));
      }
      
      // Remove chart captions that were already in the HTML to avoid duplicates
      const existingCaptions = reportClone.querySelectorAll('.text-slate-400.text-xs.mt-4');
      existingCaptions.forEach(cap => cap.remove());
      
      const content = reportClone.innerHTML;
      
      // Final cleanup and handle highlighted numbers [26]{.mark}
      const finalReportHtml = content.replace(/\[([^\]]+)\]\{\.mark\}/g, '<span class="mark">$1</span>');
      
      const blob = new Blob([header + finalReportHtml + footer], { type: 'application/msword' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${courseInfo.courseName || '课程'}_达成度分析报告.doc`;
      link.click();
      URL.revokeObjectURL(url);
      
    } catch (error) {
      console.error('Word Export Error:', error);
      setUploadError('生成 Word 报告失败。建议：1. 确保所有图表已加载完成；2. 尝试使用 Chrome 浏览器；3. 如果数据量过大，请尝试分段导出。');
      setErrorType('export');
    } finally {
      setIsProcessing(false);
      setIsExporting(false);
      setLoadingMessage('正在处理数据，请稍候...');
    }
  };

  // --- UI Components ---

  const SectionHeader = ({ icon: Icon, title, step, children }: { icon: any, title: string, step: number, children?: React.ReactNode }) => (
    <div className={cn("flex items-center justify-between mb-6 pb-4", isExporting && "mb-4 pb-0")}>
      <div className="flex items-center gap-3">
        {!isExporting && (
          <>
            <div className="flex items-center justify-center w-8 h-8 rounded-full bg-blue-600 text-white font-bold text-sm">
              {step}
            </div>
            <Icon className="w-5 h-5 text-blue-600" />
          </>
        )}
        <h2 className={cn("text-xl font-bold text-gray-800", isExporting && "text-black text-[18pt]")}>{title}</h2>
      </div>
      {!isExporting && children}
    </div>
  );

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-4 md:p-8">
      <div className="max-w-7xl mx-auto space-y-8">
        
        {/* Header */}
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
          <div id="header-title">
            <h1 className="text-3xl font-black tracking-tight text-slate-900">大学课程目标达成度分析系统</h1>
            <p className="text-slate-500 mt-1">专业、精准、可视化的课程评估工具</p>
          </div>
          <div className="flex gap-3" id="header-actions">
            <button 
              onClick={startGuide}
              className="flex items-center gap-2 px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-semibold transition-all border border-slate-200"
            >
              <HelpCircle className="w-4 h-4" />
              使用指引
            </button>
            <button 
              onClick={exportWord}
              disabled={finalData.length === 0 || isProcessing}
              className="flex items-center gap-2 px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-semibold transition-all disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-emerald-200"
            >
              <FileText className="w-4 h-4" />
              {isProcessing ? '正在导出...' : '导出 Word 报告'}
            </button>
          </div>
        </header>

        <div 
          ref={reportRef} 
          className={cn("space-y-8 bg-white p-12", isExporting && "export-mode")}
        >
          {isExporting && (
            <div className="text-center mb-12">
              <h1 className="text-3xl font-bold text-slate-900 mb-6">土木工程学院课程目标达成情况评价分析报告</h1>
            </div>
          )}

          {uploadError && (
            <div className="p-4 bg-red-50 border border-red-100 rounded-xl flex items-start gap-3 text-red-700 text-sm animate-shake" data-html2canvas-ignore>
              <AlertCircle className="w-5 h-5 mt-0.5 flex-shrink-0" />
              <div>
                <p className="font-bold">
                  {errorType === 'upload' ? '上传失败' : errorType === 'export' ? '导出失败' : '操作失败'}
                </p>
                <p className="opacity-80">{uploadError}</p>
              </div>
              <button 
                onClick={() => setUploadError(null)}
                className="ml-auto text-red-400 hover:text-red-600 transition-colors"
              >
                <Plus className="w-4 h-4 rotate-45" />
              </button>
            </div>
          )}
          {/* 栏目1：课程基本信息 */}
        <section id="section-course-info" className={cn("bg-white p-8 rounded-2xl shadow-sm border border-slate-200", isExporting && "p-0 border-none shadow-none")}>
          <SectionHeader icon={Info} title="一、课程基本信息" step={1}>
          </SectionHeader>
          {isExporting ? (
            <div className="overflow-x-auto">
              <table className="w-full border-collapse border border-slate-300 text-sm">
                <tbody>
                  <tr>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold w-32">课程名称</td>
                    <td className="border border-slate-300 p-2">{courseInfo.courseName}</td>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold w-32">课程编号</td>
                    <td className="border border-slate-300 p-2">{courseInfo.courseId}</td>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold w-32">开课学期</td>
                    <td className="border border-slate-300 p-2">{courseInfo.semester}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">课程性质</td>
                    <td className="border border-slate-300 p-2">{courseInfo.courseNature}</td>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">学 分</td>
                    <td className="border border-slate-300 p-2">{courseInfo.credits}</td>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">课内学时</td>
                    <td className="border border-slate-300 p-2">{courseInfo.classHours}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">考试/考查</td>
                    <td className="border border-slate-300 p-2">{courseInfo.examType}</td>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">开卷/闭卷</td>
                    <td className="border border-slate-300 p-2">{courseInfo.bookType}</td>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">专业班级</td>
                    <td className="border border-slate-300 p-2">{courseInfo.className}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">任课教师</td>
                    <td className="border border-slate-300 p-2">{courseInfo.teacher}</td>
                    <td className="border border-slate-300 p-2 bg-slate-50 font-bold">修读人数</td>
                    <td className="border border-slate-300 p-2" colSpan={3}>{courseInfo.studentCount}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          ) : (
            <React.Fragment>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">课程名称</label>
              <input 
                type="text" 
                value={courseInfo.courseName}
                onChange={e => setCourseInfo({...courseInfo, courseName: e.target.value})}
                placeholder="请输入课程全称"
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">课程编号</label>
              <input 
                type="text" 
                value={courseInfo.courseId}
                onChange={e => setCourseInfo({...courseInfo, courseId: e.target.value})}
                placeholder="课程编号"
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">开课学期</label>
              <input 
                type="text" 
                value={courseInfo.semester}
                onChange={e => setCourseInfo({...courseInfo, semester: e.target.value})}
                placeholder="例如: 2025-2026第一学期"
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">课程性质</label>
              <input 
                type="text" 
                value={courseInfo.courseNature}
                onChange={e => setCourseInfo({...courseInfo, courseNature: e.target.value})}
                placeholder="例如: 必修"
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">学分</label>
              <input 
                type="number" 
                step="0.1"
                value={courseInfo.credits}
                onChange={e => setCourseInfo({...courseInfo, credits: Number(e.target.value)})}
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">课内学时</label>
              <input 
                type="number" 
                value={courseInfo.classHours}
                onChange={e => setCourseInfo({...courseInfo, classHours: Number(e.target.value)})}
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">考试/考查</label>
              <input 
                type="text" 
                value={courseInfo.examType}
                onChange={e => setCourseInfo({...courseInfo, examType: e.target.value})}
                placeholder="例如: 考查"
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">开卷/闭卷</label>
              <input 
                type="text" 
                value={courseInfo.bookType}
                onChange={e => setCourseInfo({...courseInfo, bookType: e.target.value})}
                placeholder="例如: 开卷"
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">授课教师</label>
              <input 
                type="text" 
                value={courseInfo.teacher}
                onChange={e => setCourseInfo({...courseInfo, teacher: e.target.value})}
                placeholder="教师姓名"
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600">平时成绩占比 (%)</label>
              <input 
                type="number" 
                value={courseInfo.usualWeight}
                onChange={e => setCourseInfo({...courseInfo, usualWeight: Number(e.target.value), examWeight: 100 - Number(e.target.value)})}
                className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600 flex items-center gap-2">
                期末考试占比 (%)
                <span className="text-[10px] bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded">不可编辑</span>
              </label>
              <input 
                type="number" 
                value={courseInfo.examWeight}
                readOnly
                className="w-full px-4 py-2 rounded-lg border border-slate-200 bg-slate-50 text-slate-500 outline-none transition-all cursor-not-allowed"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-600 flex items-center gap-2">
                目标权重 (自动计算)
                <span className="text-[10px] bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded">不可编辑</span>
              </label>
              <div className="flex gap-2">
                {courseInfo.objectiveRatios.map((r, i) => (
                  <div 
                    key={i}
                    className="w-full px-2 py-2 rounded-lg border border-slate-200 bg-slate-50 text-slate-500 text-center font-medium"
                  >
                    {r}
                  </div>
                ))}
              </div>
            </div>

            <div className="md:col-span-3 space-y-4">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-4 h-4 text-emerald-500" />
                <h4 className="text-sm font-bold text-slate-700">平时成绩组成权重 (%)</h4>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                {courseInfo.usualAssignmentWeights.map((_, i) => (
                  <div key={i} className="relative group space-y-3 p-3 bg-white border border-slate-200 rounded-xl">
                    <button 
                      onClick={() => removeAssignment(i)}
                      className="absolute -top-2 -right-2 w-6 h-6 bg-red-500 text-white rounded-full flex items-center justify-center opacity-0 group-hover:opacity-100 transition-all shadow-sm hover:bg-red-600 z-10"
                    >
                      <Trash2 className="w-3 h-3" />
                    </button>
                    <div className="space-y-1">
                      <label className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1">
                        作业 {i + 1} 权重 (%)
                        {i === courseInfo.usualAssignmentWeights.length - 1 && courseInfo.usualAssignmentWeights.length > 1 && (
                          <span className="text-[8px] bg-slate-100 text-slate-500 px-1 rounded">自动计算</span>
                        )}
                      </label>
                      <input 
                        type="number"
                        value={courseInfo.usualAssignmentWeights[i]}
                        readOnly={i === courseInfo.usualAssignmentWeights.length - 1 && courseInfo.usualAssignmentWeights.length > 1}
                        onChange={e => {
                          const newWeights = [...courseInfo.usualAssignmentWeights];
                          const newVal = Number(e.target.value);
                          newWeights[i] = newVal;
                          
                          // Recalculate the last weight if it's not the one being changed
                          if (newWeights.length > 1 && i !== newWeights.length - 1) {
                            const sumOfOthers = newWeights.slice(0, -1).reduce((sum, w) => sum + w, 0);
                            newWeights[newWeights.length - 1] = Math.max(0, 100 - sumOfOthers);
                          }
                          
                          setCourseInfo({...courseInfo, usualAssignmentWeights: newWeights});
                        }}
                        className={cn(
                          "w-full px-3 py-1.5 text-sm rounded-lg border border-slate-100 outline-none transition-all",
                          i === courseInfo.usualAssignmentWeights.length - 1 && courseInfo.usualAssignmentWeights.length > 1
                            ? "bg-slate-100 text-slate-500 cursor-not-allowed"
                            : "bg-slate-50 focus:ring-2 focus:ring-emerald-500"
                        )}
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">对应课程目标</label>
                      <div className="flex flex-wrap gap-2">
                        {courseInfo.objectiveRatios.map((_, objIdx) => (
                          <button
                            key={objIdx}
                            onClick={() => {
                              const newObjMap = [...courseInfo.usualAssignmentObjectives];
                              const currentObjs = newObjMap[i] || [];
                              if (currentObjs.includes(objIdx)) {
                                newObjMap[i] = currentObjs.filter(o => o !== objIdx);
                              } else {
                                newObjMap[i] = [...currentObjs, objIdx];
                              }
                              setCourseInfo({...courseInfo, usualAssignmentObjectives: newObjMap});
                            }}
                            className={cn(
                              "w-6 h-6 rounded flex items-center justify-center text-[10px] font-bold transition-all",
                              (courseInfo.usualAssignmentObjectives[i] || []).includes(objIdx)
                                ? "bg-emerald-500 text-white shadow-sm"
                                : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                            )}
                          >
                            {objIdx + 1}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>
                ))}
                <button 
                  onClick={addAssignment}
                  className="flex flex-col items-center justify-center gap-2 p-3 border-2 border-dashed border-slate-200 rounded-xl text-slate-400 hover:border-blue-500 hover:text-blue-500 transition-all min-h-[120px]"
                >
                  <Plus className="w-5 h-5" />
                  <span className="text-xs font-bold">添加作业项</span>
                </button>
              </div>
            </div>

            <div className="md:col-span-3 grid grid-cols-2 md:grid-cols-4 gap-4 p-4 bg-slate-50 rounded-xl border border-dashed border-slate-300">
              <div className="flex flex-col">
                <span className="text-xs text-slate-400 uppercase font-bold tracking-wider">班级名称</span>
                <span className="text-lg font-bold text-blue-600">{courseInfo.className}</span>
              </div>
              <div className="flex flex-col">
                <span className="text-xs text-slate-400 uppercase font-bold tracking-wider">学生总人数</span>
                <span className="text-lg font-bold text-blue-600">{courseInfo.studentCount} 人</span>
              </div>
              </div>
            </div>
          </React.Fragment>
        )}
      </section>

        {/* 栏目2：课程目标 */}
        <section id="section-objectives" className={cn("bg-white p-8 rounded-2xl shadow-sm border border-slate-200", isExporting && "p-0 border-none shadow-none")}>
          <SectionHeader icon={TrendingUp} title="二、课程目标" step={2} />
          <div className="space-y-4">
            {isExporting ? (
              <div className="overflow-x-auto">
                <table className="w-full border-collapse border border-slate-300 text-sm">
                  <tbody>
                    {courseInfo.objectiveDescriptions.map((desc, i) => (
                      <tr key={i}>
                        <td className="border border-slate-300 p-4 bg-slate-50 font-bold w-32 text-center">课程目标{i+1}</td>
                        <td className="border border-slate-300 p-4 leading-relaxed" style={{ textAlign: 'left', textIndent: '2em', whiteSpace: 'pre-wrap' }}>{desc}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="space-y-4">
                {courseInfo.objectiveDescriptions.map((desc, i) => (
                  <div key={i} className="space-y-1 group relative">
                    <div className="flex justify-between items-center">
                      <label className="text-xs font-bold text-slate-500">课程目标 {i + 1}</label>
                      {courseInfo.objectiveDescriptions.length > 1 && (
                        <button 
                          onClick={() => {
                            const newDescs = courseInfo.objectiveDescriptions.filter((_, idx) => idx !== i);
                            setCourseInfo({...courseInfo, objectiveDescriptions: newDescs, objectiveRatios: new Array(newDescs.length).fill(0)});
                          }}
                          className="opacity-0 group-hover:opacity-100 text-red-400 hover:text-red-600 transition-all"
                          title="删除此目标"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                    <textarea 
                      value={desc}
                      onChange={e => {
                        const newDescs = [...courseInfo.objectiveDescriptions];
                        newDescs[i] = e.target.value;
                        setCourseInfo({...courseInfo, objectiveDescriptions: newDescs});
                      }}
                      className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all min-h-[80px]"
                    />
                  </div>
                ))}
                <div className="flex justify-center pt-2">
                  <button 
                    onClick={() => {
                      setCourseInfo(prev => ({
                        ...prev,
                        objectiveDescriptions: [...prev.objectiveDescriptions, ''],
                        objectiveRatios: [...prev.objectiveRatios, 0]
                      }));
                    }}
                    className="flex items-center gap-2 px-4 py-2 bg-slate-50 text-slate-600 rounded-xl hover:bg-slate-100 transition-all border border-slate-200 text-sm font-bold"
                  >
                    <Plus className="w-4 h-4" /> 添加课程目标
                  </button>
                </div>
              </div>
            )}
          </div>
        </section>

        {isExporting && (
          <section className="bg-white p-0 rounded-none border-none shadow-none break-before-page">
            <SectionHeader icon={TrendingUp} title={isExporting ? "三、课程考核方式与成绩评定" : "三.一、课程考核方式与成绩评定"} step={3} />
            <div className="overflow-x-auto">
              <table className="w-full border-collapse border border-slate-300 text-sm text-center">
                <thead style={{ msoHeaderRow: 'yes' }}>
                  <tr className="bg-slate-50" style={{ msoHeaderRow: 'yes' }}>
                    <th className="border border-slate-300 p-2">考核环节</th>
                    {courseInfo.objectiveRatios.map((_, i) => (
                      <th key={i} className="border border-slate-300 p-2">课程目标{i + 1}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {courseInfo.usualAssignmentWeights.map((weight, i) => (
                    <tr key={i}>
                      <td className="border border-slate-300 p-2">平时考核{String.fromCharCode(65 + i)}（{weight}%）</td>
                      {courseInfo.objectiveRatios.map((_, objIdx) => {
                        const isMapped = courseInfo.usualAssignmentObjectives[i].includes(objIdx);
                        if (!isMapped) return <td key={objIdx} className="border border-slate-300 p-2">-</td>;
                        
                        // Calculate weight within the objective
                        // This is tricky. The reference shows how much of the objective is covered by this assignment.
                        // If an assignment covers multiple objectives, we assume it's split or fully covers?
                        // Usually it's the weight of the assignment relative to the objective's total weight.
                        return <td key={objIdx} className="border border-slate-300 p-2">{(weight * courseInfo.usualWeight / 100).toFixed(1)}%</td>;
                      })}
                    </tr>
                  ))}
                  <tr>
                    <td className="border border-slate-300 p-2">期末考试（{courseInfo.examWeight}%）</td>
                    {courseInfo.objectiveRatios.map((_, objIdx) => {
                      // Find questions mapped to this objective
                      const objExamMax = questionMappings
                        .filter(m => m.objectiveIndex === objIdx)
                        .reduce((sum, m) => sum + m.maxScore, 0);
                      const totalExamMax = questionMappings.reduce((sum, m) => sum + m.maxScore, 0);
                      const examContribution = totalExamMax > 0 ? (objExamMax / totalExamMax) * courseInfo.examWeight : 0;
                      
                      return <td key={objIdx} className="border border-slate-300 p-2">{examContribution.toFixed(1)}%</td>;
                    })}
                  </tr>
                  <tr className="font-bold bg-slate-50">
                    <td className="border border-slate-300 p-2">合计</td>
                    {courseInfo.objectiveRatios.map((r, i) => (
                      <td key={i} className="border border-slate-300 p-2">{r}%</td>
                    ))}
                  </tr>
                  <tr className="bg-slate-50">
                    <td className="border border-slate-300 p-2">总成绩</td>
                    <td className="border border-slate-300 p-2" colSpan={courseInfo.objectiveRatios.length}>
                      平时成绩 × {courseInfo.usualWeight}% + 期末考试 × {courseInfo.examWeight}%
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>
        )}

        {isExporting && stats && (
          <section className="bg-white p-0 rounded-none border-none shadow-none break-before-page">
            <SectionHeader icon={FileText} title="四、课程目标达成原始记录表" step={4} />
            
            <div className="overflow-x-auto">
              <table className="w-full border-collapse border border-black text-[9px]">
                <thead style={{ msoHeaderRow: 'yes' }}>
                  <tr style={{ msoHeaderRow: 'yes' }}>
                    <th colSpan={3} className="border border-black p-1 bg-slate-50">课程目标</th>
                    {courseInfo.objectiveRatios.map((_, i) => (
                      <th key={i} colSpan={3} className="border border-black p-1">课程目标{i + 1}</th>
                    ))}
                    <th className="border border-black p-1">期末卷面成绩</th>
                    <th className="border border-black p-1">课程目标评价值</th>
                  </tr>
                  <tr style={{ msoHeaderRow: 'yes' }}>
                    <th colSpan={3} className="border border-black p-1 bg-slate-50">考核方式</th>
                    {courseInfo.objectiveRatios.map((_, i) => (
                      <React.Fragment key={i}>
                        <th className="border border-black p-1">期末考试</th>
                        <th className="border border-black p-1">平时考核</th>
                        <th className="border border-black p-1">评价值</th>
                      </React.Fragment>
                    ))}
                    <th className="border border-black p-1"></th>
                    <th className="border border-black p-1"></th>
                  </tr>
                  <tr style={{ msoHeaderRow: 'yes' }}>
                    <th colSpan={3} className="border border-black p-1 bg-slate-50">满分</th>
                    {courseInfo.objectiveRatios.map((ratio, i) => (
                      <React.Fragment key={i}>
                        <th className="border border-black p-1">{ratio}</th>
                        <th className="border border-black p-1">100</th>
                        <th className="border border-black p-1">100</th>
                      </React.Fragment>
                    ))}
                    <th className="border border-black p-1">100</th>
                    <th className="border border-black p-1">100</th>
                  </tr>
                  <tr className="bg-slate-50 font-bold" style={{ msoHeaderRow: 'yes' }}>
                    <th className="border border-black p-1">学号</th>
                    <th className="border border-black p-1">姓名</th>
                    <th className="border border-black p-1">班级</th>
                    {courseInfo.objectiveRatios.map((_, i) => (
                      <React.Fragment key={i}>
                        <th className="border border-black p-1">T{String.fromCharCode(65 + i)}1</th>
                        <th className="border border-black p-1">T{String.fromCharCode(65 + i)}2</th>
                        <th className="border border-black p-1">T{String.fromCharCode(65 + i)}A</th>
                      </React.Fragment>
                    ))}
                    <th className="border border-black p-1">TTT</th>
                    <th className="border border-black p-1">TTA</th>
                  </tr>
                </thead>
                <tbody>
                  {finalData.map((student, idx) => (
                    <tr key={idx}>
                      <td className="border border-black p-1 text-center">{student.id}</td>
                      <td className="border border-black p-1 text-center">{student.name}</td>
                      <td className="border border-black p-1 text-center">{student.className || courseInfo.className}</td>
                      {student.objectiveScores.map((obj, i) => (
                        <React.Fragment key={i}>
                          <td className="border border-black p-1 text-center">{obj.examScore.toFixed(1)}</td>
                          <td className="border border-black p-1 text-center">{obj.usualScore.toFixed(1)}</td>
                          <td className="border border-black p-1 text-center">{Math.round(obj.achievement * 100)}</td>
                        </React.Fragment>
                      ))}
                      <td className="border border-black p-1 text-center">{student.totalExamScore.toFixed(1)}</td>
                      <td className="border border-black p-1 text-center font-bold">{Math.round(student.finalGrade)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {/* 栏目2：成绩数据输入 */}
        <section id="section-upload" className="bg-white p-8 rounded-2xl shadow-sm border border-slate-200" data-html2canvas-ignore>
          <SectionHeader icon={Upload} title="成绩数据输入" step={2} />
          
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
            {/* Exam Structure Upload */}
            <div className="relative group" data-html2canvas-ignore>
              <div className="absolute -inset-1 bg-gradient-to-r from-blue-600 to-indigo-600 rounded-2xl blur opacity-10 group-hover:opacity-20 transition duration-1000 group-hover:duration-200"></div>
              <div className="relative flex flex-col items-center justify-center p-8 border-2 border-dashed border-slate-200 rounded-2xl bg-white hover:border-blue-500 transition-all cursor-pointer">
                <Table className="w-10 h-10 text-blue-500 mb-4" />
                <h3 className="font-bold text-slate-800">上传考试结构表</h3>
                <p className="text-xs text-slate-400 mt-2 text-center">支持 .xlsx, .xls 格式<br/>包含题号、满分、对应课程目标<br/>(题号和阅卷系统下载小题得分表的题号一致)</p>
                <input type="file" accept=".xlsx, .xls" onChange={handleExamStructureUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
                
                {questionMappings.length > 0 && (
                  <div className="mt-4 flex items-center gap-2 text-green-600 text-sm font-bold">
                    <CheckCircle2 className="w-4 h-4" /> 已解析 {questionMappings.length} 个小题结构
                  </div>
                )}
              </div>
            </div>

            {/* Exam Detail Upload */}
            <div className="relative group" data-html2canvas-ignore>
              <div className="absolute -inset-1 bg-gradient-to-r from-indigo-600 to-purple-600 rounded-2xl blur opacity-10 group-hover:opacity-20 transition duration-1000 group-hover:duration-200"></div>
              <div className="relative flex flex-col items-center justify-center p-8 border-2 border-dashed border-slate-200 rounded-2xl bg-white hover:border-indigo-500 transition-all cursor-pointer">
                <BarChart3 className="w-10 h-10 text-indigo-500 mb-4" />
                <h3 className="font-bold text-slate-800">上传考试得分明细表</h3>
                <p className="text-xs text-slate-400 mt-2 text-center">支持 .xlsx, .xls 格式<br/>包含学号、姓名、各小题得分<br/>阅卷系统下载的小题得分表可直接上传</p>
                <input type="file" accept=".xlsx, .xls" onChange={handleExamDetailUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
                
                {students.length > 0 && (
                  <div className="mt-4 flex items-center gap-2 text-green-600 text-sm font-bold">
                    <CheckCircle2 className="w-4 h-4" /> 已解析 {students.length} 名学生得分
                  </div>
                )}
              </div>
            </div>

            {/* Usual Upload */}
            <div className="relative group" data-html2canvas-ignore>
              <div className="absolute -inset-1 bg-gradient-to-r from-emerald-600 to-teal-600 rounded-2xl blur opacity-10 group-hover:opacity-20 transition duration-1000 group-hover:duration-200"></div>
              <div className="relative flex flex-col items-center justify-center p-8 border-2 border-dashed border-slate-200 rounded-2xl bg-white hover:border-emerald-500 transition-all cursor-pointer">
                <Upload className="w-10 h-10 text-emerald-500 mb-4" />
                <h3 className="font-bold text-slate-800">上传平时成绩表</h3>
                <p className="text-xs text-slate-400 mt-2 text-center">支持 .xlsx, .xls 格式<br/>包含学号、平时总评成绩<br/>(按模板填写，包含每次作业得分和平时成绩总分)</p>
                <input type="file" accept=".xlsx, .xls" onChange={handleUsualUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
                
                {students.some(s => s.usualTotal > 0) && (
                  <div className="mt-4 flex items-center gap-2 text-green-600 text-sm font-bold">
                    <CheckCircle2 className="w-4 h-4" /> 平时成绩已匹配
                  </div>
                )}

                {students.length > 0 && !students.some(s => s.usualTotal > 0) && (
                  <button 
                    onClick={(e) => {
                      e.stopPropagation();
                      downloadUsualTemplate();
                    }}
                    className="mt-6 px-4 py-2 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-lg text-xs font-bold flex items-center gap-2 hover:bg-emerald-100 transition-colors relative z-10"
                  >
                    <FileDown className="w-4 h-4" />
                    下载平时成绩模板
                  </button>
                )}
              </div>
            </div>
          </div>

          {students.length > 0 && (
            <div className="mt-8 flex flex-col gap-4">
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-600">
                    <CheckCircle2 className="w-6 h-6" />
                  </div>
                  <div>
                    <p className="font-bold text-slate-800">数据加载成功</p>
                    <p className="text-xs text-slate-500">已加载 {students.length} 名学生的数据</p>
                  </div>
                </div>
                <button 
                  onClick={() => setShowDataTable(!showDataTable)}
                  data-html2canvas-ignore
                  className="px-4 py-2 bg-white border border-slate-200 rounded-lg text-sm font-bold text-slate-600 hover:bg-slate-50 transition-colors flex items-center gap-2"
                >
                  <Table className="w-4 h-4" />
                  {showDataTable ? '隐藏数据详情' : '查看数据详情'}
                </button>
              </div>

              {showDataTable && (
                <motion.div 
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  className="overflow-x-auto rounded-xl border border-slate-200"
                >
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200">
                        <th className="p-3 font-bold text-slate-500">学号</th>
                        <th className="p-3 font-bold text-slate-500">姓名</th>
                        <th className="p-3 font-bold text-slate-500">班级</th>
                        <th className="p-3 font-bold text-slate-500">考试总分</th>
                        <th className="p-3 font-bold text-slate-500">平时成绩</th>
                      </tr>
                    </thead>
                    <tbody>
                      {students.slice(0, 10).map((s, i) => (
                        <tr key={i} className="border-b border-slate-100 hover:bg-slate-50/50">
                          <td className="p-3 font-mono">{s.id}</td>
                          <td className="p-3">{s.name}</td>
                          <td className="p-3">{s.className}</td>
                          <td className="p-3 font-bold text-blue-600">{s.examTotal}</td>
                          <td className="p-3 font-bold text-emerald-600">{s.usualTotal}</td>
                        </tr>
                      ))}
                      {students.length > 10 && (
                        <tr className="bg-slate-50/30">
                          <td colSpan={5} className="p-3 text-center text-slate-400 italic">
                            ... 仅显示前 10 条数据，共 {students.length} 条 ...
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </motion.div>
              )}
            </div>
          )}

          {/* Question Mapping Configuration */}
          {questionMappings.length > 0 && (
            <div className="mt-8 p-6 bg-slate-50 rounded-xl border border-slate-200">
              <div className="flex items-center justify-between mb-4">
                <h3 className="font-bold text-slate-800 flex items-center gap-2">
                  <TrendingUp className="w-4 h-4 text-blue-500" />
                  题目与课程目标映射配置
                </h3>
                <span className="text-xs text-slate-400">点击目标编号可切换 (1-4)</span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-8 gap-3">
                {questionMappings.map((m, idx) => (
                  <div key={m.questionId} className="flex flex-col p-2 bg-white rounded-lg border border-slate-200 shadow-sm">
                    <span className="text-[10px] font-bold text-slate-400 truncate">{m.questionId}</span>
                    <div className="flex items-center justify-between mt-1">
                      <input 
                        type="number" 
                        value={m.maxScore}
                        onChange={e => {
                          const newMappings = [...questionMappings];
                          newMappings[idx].maxScore = Number(e.target.value);
                          setQuestionMappings(newMappings);
                        }}
                        className="w-10 text-xs font-bold text-blue-600 outline-none"
                        title="满分"
                      />
                      <button 
                        onClick={() => {
                          const newMappings = [...questionMappings];
                          newMappings[idx].objectiveIndex = (newMappings[idx].objectiveIndex + 1) % courseInfo.objectiveRatios.length;
                          setQuestionMappings(newMappings);
                        }}
                        className={cn(
                          "w-6 h-6 rounded-full text-[10px] font-bold flex items-center justify-center transition-colors",
                          m.objectiveIndex === 0 ? "bg-blue-100 text-blue-600" :
                          m.objectiveIndex === 1 ? "bg-emerald-100 text-emerald-600" :
                          m.objectiveIndex === 2 ? "bg-amber-100 text-amber-600" :
                          m.objectiveIndex === 3 ? "bg-purple-100 text-purple-600" :
                          "bg-rose-100 text-rose-600"
                        )}
                      >
                        目{m.objectiveIndex + 1}
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </section>

        {/* 栏目3：考核成绩统计结果 */}
        <AnimatePresence>
          {stats && !isExporting && (
            <motion.section 
              id="section-analysis"
              initial={isExporting ? false : { opacity: 0, y: 20 }}
              animate={isExporting ? false : { opacity: 1, y: 0 }}
              className={cn("bg-white p-8 rounded-2xl shadow-sm border border-slate-200", isExporting && "p-0 border-none shadow-none break-before-page")}
            >
              <div className="flex justify-between items-center mb-8">
                <SectionHeader icon={FileText} title={isExporting ? "三、考核成绩统计结果" : "考核成绩统计结果"} step={3} />
                {!isExporting && (
                  <button
                    onClick={exportToExcel}
                    className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-xl hover:bg-emerald-700 transition-all shadow-sm font-bold text-sm"
                  >
                    <FileSpreadsheet className="w-4 h-4" />
                    下载考核成绩统计表 (Excel)
                  </button>
                )}
              </div>
              <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-8">
                {[
                  { label: '最高分', value: stats.max, color: 'text-blue-600' },
                  { label: '最低分', value: stats.min, color: 'text-slate-600' },
                  { label: '平均分', value: stats.avg, color: 'text-blue-600' },
                  { label: '中位数', value: stats.median, color: 'text-slate-600' },
                  { label: '及格率', value: `${Math.round((stats.passCount / finalData.length) * 100)}%`, color: 'text-emerald-600' },
                ].map((item, i) => (
                  <div key={i} className="p-4 bg-slate-50 rounded-xl border border-slate-100 text-center">
                    <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">{item.label}</p>
                    <p className={cn("text-2xl font-black", item.color)}>{item.value}</p>
                  </div>
                ))}
              </div>
              
              <div className="overflow-hidden rounded-xl border border-slate-200">
                <table className="w-full text-left border-collapse">
                  <thead style={{ msoHeaderRow: 'yes' }}>
                    <tr className="bg-slate-50" style={{ msoHeaderRow: 'yes' }}>
                      <th className="p-4 text-xs font-bold text-slate-500 uppercase">分数段</th>
                      <th className="p-4 text-xs font-bold text-slate-500 uppercase">优秀 (≥85)</th>
                      <th className="p-4 text-xs font-bold text-slate-500 uppercase">良好 (70-84)</th>
                      <th className="p-4 text-xs font-bold text-slate-500 uppercase">中等 (60-69)</th>
                      <th className="p-4 text-xs font-bold text-slate-500 uppercase">不及格 (&lt;60)</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="border-t border-slate-100 hover:bg-slate-50/50 transition-colors">
                      <td className="p-4 font-bold text-slate-400">人数</td>
                      <td className="p-4 font-black text-emerald-600 text-lg">{stats.grades['优秀']}</td>
                      <td className="p-4 font-black text-blue-600 text-lg">{stats.grades['良好']}</td>
                      <td className="p-4 font-black text-amber-600 text-lg">{stats.grades['中等']}</td>
                      <td className="p-4 font-black text-red-600 text-lg">{stats.grades['不及格']}</td>
                    </tr>
                    <tr className="border-t border-slate-100 hover:bg-slate-50/50 transition-colors">
                      <td className="p-4 font-bold text-slate-400">占比</td>
                      <td className="p-4 text-sm font-semibold text-slate-600">{Math.round((stats.grades['优秀'] / finalData.length) * 100)}%</td>
                      <td className="p-4 text-sm font-semibold text-slate-600">{Math.round((stats.grades['良好'] / finalData.length) * 100)}%</td>
                      <td className="p-4 text-sm font-semibold text-slate-600">{Math.round((stats.grades['中等'] / finalData.length) * 100)}%</td>
                      <td className="p-4 text-sm font-semibold text-slate-600">{Math.round((stats.grades['不及格'] / finalData.length) * 100)}%</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        {/* 栏目4：可视化分析面板 */}
        <AnimatePresence>
          {stats && !isExporting && (
            <div className="space-y-8">
              <motion.section 
                initial={isExporting ? false : { opacity: 0, y: 20 }}
                animate={isExporting ? false : { opacity: 1, y: 0 }}
                className={cn("bg-white p-8 rounded-2xl shadow-sm border border-slate-200", isExporting && "p-0 border-none shadow-none")}
              >
                <SectionHeader icon={BarChart3} title={isExporting ? "五、可视化分析面板" : "可视化分析面板"} step={isExporting ? 5 : 4} />
                <div className="p-6 bg-blue-50 rounded-2xl border border-blue-100 mb-4">
                  <p className="text-sm text-blue-700 leading-relaxed">
                    本部分通过直方图、饼图及散点图，直观展示班级整体成绩分布情况及平时与考试成绩的相关性。
                  </p>
                </div>
              </motion.section>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* 综合成绩分布直方图 */}
                <motion.section 
                  initial={isExporting ? false : { opacity: 0, scale: 0.95 }}
                  animate={isExporting ? false : { opacity: 1, scale: 1 }}
                  className="p-8 bg-white rounded-2xl shadow-sm border border-slate-200"
                >
                  <div className="flex justify-between items-center mb-6">
                    <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                      <div className="w-1.5 h-5 bg-blue-600 rounded-full"></div>
                      综合成绩分布直方图
                    </h3>
                    {!isExporting && (
                      <button
                        onClick={() => downloadChart('chart-histogram', '综合成绩分布直方图')}
                        className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                        title="下载图片"
                      >
                        <Download className="w-4 h-4" />
                      </button>
                    )}
                  </div>
                  <div id="chart-histogram" className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart 
                        data={[
                          { range: '0-59', count: stats.grades['不及格'] },
                          { range: '60-69', count: stats.grades['中等'] },
                          { range: '70-84', count: stats.grades['良好'] },
                          { range: '85-100', count: stats.grades['优秀'] },
                        ]}
                        margin={isExporting ? { top: 40, right: 40, left: 40, bottom: 40 } : { top: 20, right: 30, left: 20, bottom: 5 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis 
                          dataKey="range" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ 
                            fontSize: 12, 
                            fill: '#64748b',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                        <YAxis 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ 
                            fontSize: 12, 
                            fill: '#64748b',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                        <Tooltip 
                          cursor={{ fill: '#f1f5f9' }}
                          contentStyle={{ 
                            borderRadius: '12px', 
                            border: 'none', 
                            boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }}
                        />
                        <Bar dataKey="count" radius={[6, 6, 0, 0]} isAnimationActive={!isExporting}>
                          { [0, 1, 2, 3].map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={Object.values(GRADE_COLORS)[index]} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </motion.section>

                {/* 各分数段占比饼图 */}
                <motion.section 
                  initial={isExporting ? false : { opacity: 0, scale: 0.95 }}
                  animate={isExporting ? false : { opacity: 1, scale: 1 }}
                  className="p-8 bg-white rounded-2xl shadow-sm border border-slate-200"
                >
                  <div className="flex justify-between items-center mb-6">
                    <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                      <div className="w-1.5 h-5 bg-emerald-600 rounded-full"></div>
                      各分数段占比饼图
                    </h3>
                    {!isExporting && (
                      <button
                        onClick={() => downloadChart('chart-pie', '各分数段占比饼图')}
                        className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg transition-all"
                        title="下载图片"
                      >
                        <Download className="w-4 h-4" />
                      </button>
                    )}
                  </div>
                  <div id="chart-pie" className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart margin={isExporting ? { top: 20, right: 20, bottom: 20, left: 20 } : { top: 0, right: 0, bottom: 0, left: 0 }}>
                        <Pie
                          data={Object.entries(stats.grades).map(([name, value]) => ({ name, value }))}
                          cx="50%"
                          cy="50%"
                          innerRadius={60}
                          outerRadius={100}
                          paddingAngle={5}
                          dataKey="value"
                          isAnimationActive={!isExporting}
                        >
                          {Object.entries(stats.grades).map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={Object.values(GRADE_COLORS)[index]} />
                          ))}
                        </Pie>
                        <Tooltip 
                          contentStyle={{ 
                            borderRadius: '12px', 
                            border: 'none', 
                            boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                        <Legend 
                          verticalAlign="bottom" 
                          height={36}
                          wrapperStyle={{ fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif" }}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </motion.section>

                {/* 平时成绩 vs 考试成绩散点图 */}
                <motion.section 
                  initial={isExporting ? false : { opacity: 0, scale: 0.95 }}
                  animate={isExporting ? false : { opacity: 1, scale: 1 }}
                  className="lg:col-span-2 p-8 bg-white rounded-2xl shadow-sm border border-slate-200"
                >
                  <div className="flex justify-between items-center mb-6">
                    <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                      <div className="w-1.5 h-5 bg-amber-600 rounded-full"></div>
                      平时成绩 vs 考试成绩散点图 (样本量: {finalData.length}人)
                    </h3>
                    {!isExporting && (
                      <button
                        onClick={() => downloadChart('chart-scatter', '平时成绩vs考试成绩散点图')}
                        className="p-1.5 text-slate-400 hover:text-amber-600 hover:bg-amber-50 rounded-lg transition-all"
                        title="下载图片"
                      >
                        <Download className="w-4 h-4" />
                      </button>
                    )}
                  </div>
                  <div id="chart-scatter" className="h-[400px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ScatterChart 
                        margin={isExporting ? { top: 40, right: 40, bottom: 40, left: 40 } : { top: 20, right: 20, bottom: 20, left: 20 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                        <XAxis 
                          type="number" 
                          dataKey="x" 
                          name="平时成绩" 
                          unit="分" 
                          domain={[0, 100]} 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ 
                            fontSize: 12, 
                            fill: '#64748b',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }}
                        />
                        <YAxis 
                          type="number" 
                          dataKey="y" 
                          name="考试成绩" 
                          unit="分" 
                          domain={[0, 100]} 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ 
                            fontSize: 12, 
                            fill: '#64748b',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }}
                        />
                        <ZAxis type="number" range={[50, 50]} />
                        <Tooltip 
                          cursor={{ strokeDasharray: '3 3' }} 
                          contentStyle={{ 
                            borderRadius: '12px', 
                            border: 'none', 
                            boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }}
                        />
                        <Scatter name="学生成绩" data={finalData.map(s => ({ x: s.usualTotal, y: s.examTotal, name: s.name }))} fill="#3b82f6" isAnimationActive={!isExporting} />
                        <ReferenceLine 
                          x={60} 
                          stroke="#ef4444" 
                          strokeDasharray="3 3" 
                          label={{ 
                            position: 'top', 
                            value: '及格线', 
                            fill: '#ef4444', 
                            fontSize: 10,
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                        <ReferenceLine 
                          y={60} 
                          stroke="#ef4444" 
                          strokeDasharray="3 3" 
                          label={{ 
                            position: 'right', 
                            value: '及格线', 
                            fill: '#ef4444', 
                            fontSize: 10,
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>
                </motion.section>
              </div>
            </div>
          )}
        </AnimatePresence>

        {/* 栏目6：持续改进措施编辑 */}
        {!isExporting && stats && (
          <section className="bg-white p-8 rounded-2xl shadow-sm border border-slate-200" data-html2canvas-ignore>
            <SectionHeader icon={FileText} title="持续改进措施编辑" step={6} />
            <div className="space-y-6">
              <div className="space-y-2">
                <div className="flex justify-between items-center">
                  <label className="text-sm font-semibold text-slate-600">课程目标达成情况评价</label>
                  <button 
                    onClick={() => {
                      const overallAvg = Math.round((objectiveStats.reduce((sum, obj) => sum + obj.avg, 0) / objectiveStats.length) * 100) / 100;
                      let text = `本次课程目标达成评价值平均为 ${overallAvg}。`;
                      const achievedCount = objectiveStats.filter(obj => obj.avg >= 0.65).length;
                      if (achievedCount === objectiveStats.length) text += `各课程目标均达到了预期教学效果（达成度均 ≥ 0.65）。`;
                      else text += `共有 ${achievedCount} 个课程目标达到预期教学效果。`;
                      text += `从各课程目标达成度情况来看：`;
                      objectiveStats.forEach((obj, idx) => {
                        text += `课程目标 ${obj.index} 的达成度平均值为 ${obj.avg}（${obj.conclusion}）${idx === objectiveStats.length - 1 ? '。' : '；'}`;
                      });
                      const sortedStats = [...objectiveStats].sort((a, b) => a.avg - b.avg);
                      const lowest = sortedStats[0];
                      const highest = sortedStats[sortedStats.length - 1];
                      text += `其中，课程目标 ${highest.index} 达成情况最好。`;
                      if (lowest.avg < 0.7) text += `课程目标 ${lowest.index} 的达成度相对较低（${lowest.avg}），在今后的教学中需进一步加强对该目标对应知识点的讲解与训练。`;
                      
                      // Generate dynamic improvement measures
                      const improvementMeasures: string[] = [];
                      objectiveStats.forEach(obj => {
                        const { index, avg, conclusion } = obj;
                        const desc = courseInfo.objectiveDescriptions[index - 1] || '';
                        const templates = {
                          '优秀达成': [
                            `针对课程目标${index}（${desc.substring(0, 15)}...），学生掌握情况极佳。后续将引入更具挑战性的案例分析，进一步拓宽学生的专业视野。`,
                            `课程目标${index}达成度处于高位。计划在下一轮教学中增加前沿技术讲座，保持学生对该领域知识的探索热情。`,
                            `鉴于课程目标${index}的优秀表现，将总结现有教学经验，并在其他相关章节推广启发式教学法。`,
                            `课程目标${index}达成效果显著，学生对核心概念掌握扎实。未来将探索跨学科协作项目，提升综合应用能力。`
                          ],
                          '良好达成': [
                            `课程目标${index}达成情况良好。未来将加强课堂互动，通过更多的小组讨论深化学生对复杂概念的理解。`,
                            `针对课程目标${index}，计划优化作业设计，增加实践性环节，以巩固学生对理论知识的应用能力。`,
                            `课程目标${index}表现稳健。将进一步完善教学大纲，确保知识点的衔接更加紧密。`,
                            `该目标达成度较为理想。计划在后续课程中引入更多工程实际案例，强化理论联系实际的教学导向。`
                          ],
                          '基本达成': [
                            `课程目标${index}仅为基本达成，反映出部分学生在“${desc.substring(0, 10)}”方面存在薄弱环节。需增加课后辅导频率。`,
                            `针对课程目标${index}达成度偏低的问题，计划调整教学重点，增加基础知识的复习课时，并强化过程化考核。`,
                            `课程目标${index}的达成度有待提高。将重新设计实验环节，提高学生的动手能力和解决实际问题的能力。`,
                            `分析发现学生在目标${index}相关的复杂计算上存在短板。未来将通过增加课堂练习和专项测试来精准突破。`
                          ],
                          '未达成': [
                            `课程目标${index}未达成，形势严峻。必须重新审视该部分的教学设计，分析学生普遍困惑的原因，并进行专题补课。`,
                            `针对课程目标${index}未达标的情况，将组织任课教师集体研讨，改进教学手段，并对不及格学生进行重点帮扶。`,
                            `课程目标${index}达成度极低。计划在下一学期大幅增加该部分的学时分配，并引入更多直观的教学辅助工具。`,
                            `目标${index}的达成度未达预期。初步分析是由于前置知识衔接不畅，计划在开课初期增加前置知识的摸底与补强。`
                          ]
                        };
                        const category = conclusion.includes('优秀') ? '优秀达成' : 
                                         conclusion.includes('良好') ? '良好达成' : 
                                         conclusion.includes('基本') ? '基本达成' : '未达成';
                        const options = templates[category] || templates['基本达成'];
                        const selected = options[Math.floor(Math.random() * options.length)];
                        improvementMeasures.push(`${index}. ${selected}`);
                      });
                      
                      setEvaluationMeasures(prev => ({ 
                        ...prev, 
                        evaluation: text,
                        improvement: improvementMeasures.join('\n')
                      }));
                    }}
                    className="text-xs text-blue-600 hover:text-blue-700 font-bold flex items-center gap-1"
                  >
                    <TrendingUp className="w-3 h-3" />
                    根据分析结果重新生成
                  </button>
                </div>
                <textarea 
                  value={evaluationMeasures.evaluation}
                  onChange={e => setEvaluationMeasures({...evaluationMeasures, evaluation: e.target.value})}
                  className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all min-h-[120px]"
                />
              </div>
              <div className="space-y-2">
                <label className="text-sm font-semibold text-slate-600">持续改进措施</label>
                <textarea 
                  value={evaluationMeasures.improvement}
                  onChange={e => setEvaluationMeasures({...evaluationMeasures, improvement: e.target.value})}
                  className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all min-h-[120px]"
                />
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                  <label className="text-sm font-semibold text-slate-600">课程责任教授意见</label>
                  <input 
                    type="text" 
                    value={evaluationMeasures.professorOpinion}
                    onChange={e => setEvaluationMeasures({...evaluationMeasures, professorOpinion: e.target.value})}
                    className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-sm font-semibold text-slate-600">评价时间</label>
                  <input 
                    type="text" 
                    value={evaluationMeasures.evaluationDate}
                    onChange={e => setEvaluationMeasures({...evaluationMeasures, evaluationDate: e.target.value})}
                    className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                  />
                </div>
              </div>
              <div className="space-y-2">
                <label className="text-sm font-semibold text-slate-600">审核意见</label>
                <textarea 
                  value={evaluationMeasures.auditOpinion}
                  onChange={e => setEvaluationMeasures({...evaluationMeasures, auditOpinion: e.target.value})}
                  className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all min-h-[80px]"
                />
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                  <label className="text-sm font-semibold text-slate-600">审核日期</label>
                  <input 
                    type="text" 
                    value={evaluationMeasures.auditDate}
                    onChange={e => setEvaluationMeasures({...evaluationMeasures, auditDate: e.target.value})}
                    className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                  />
                </div>
              </div>
            </div>
          </section>
        )}
        <AnimatePresence>
          {objectiveStats.length > 0 && (
            <div className="space-y-8">
              <motion.section 
                initial={isExporting ? false : { opacity: 0, y: 20 }}
                animate={isExporting ? false : { opacity: 1, y: 0 }}
                className={cn("bg-white p-8 rounded-2xl shadow-sm border border-slate-200", isExporting && "p-0 border-none shadow-none")}
              >
                <SectionHeader icon={TrendingUp} title={isExporting ? "五、课程目标达成情况分析" : "五、课程目标达成情况分析"} step={5} />
                <div className="p-6 bg-blue-50 rounded-2xl border border-blue-100 mb-4">
                  <p className="text-sm text-blue-700 leading-relaxed">
                    本部分对课程设置的 {objectiveStats.length} 个目标进行详细达成度分析。分析基于各目标对应的考核小题得分率，结合平时成绩与考试成绩的加权计算得出。
                  </p>
                </div>
              </motion.section>

              {objectiveStats.map((obj, i) => (
                <motion.section 
                  key={i}
                  initial={isExporting ? false : { opacity: 0, y: 20 }}
                  animate={isExporting ? false : { opacity: 1, y: 0 }}
                  transition={isExporting ? undefined : { delay: i * 0.1 }}
                  className={cn("bg-white p-8 rounded-2xl shadow-sm border border-slate-200", isExporting && "p-0 border-none shadow-none break-before-page")}
                >
                  <div className="flex justify-between items-start mb-6">
                    <div>
                      <h3 className="text-xl font-bold text-slate-800">{i + 1}. 课程目标 {obj.index}</h3>
                      <div className="mt-4 space-y-4">
                        <div className="p-4 bg-slate-50 rounded-xl border border-slate-100">
                          {!isExporting && (
                            <p className="text-slate-700 text-sm leading-relaxed font-bold mb-2">
                              目标内容：<span className="font-normal">{courseInfo.objectiveDescriptions[i]}</span>
                            </p>
                          )}
                          <p className="text-slate-600 text-sm leading-relaxed text-justify">
                            根据《{courseInfo.courseName}》课程目标{obj.index}的学生达成情况分布图（图{obj.index}）显示，本班级学生的达成评价值呈现出明显的分布特征。
                            经统计分析，全班共有 {finalData.length} 名学生参与考核，其中达成度小于 0.65（未达成）的学生人数为 <span className="font-bold text-red-600">{obj.lowCount}</span> 人，
                            占比约为 {Math.round((obj.lowCount / finalData.length) * 100)}%；达成度大于等于 0.8（达成优秀）的学生人数为 <span className="font-bold text-emerald-600">{obj.highCount}</span> 人，
                            占比约为 {Math.round((obj.highCount / finalData.length) * 100)}%。
                            该目标的平均达成度为 <span className="font-bold text-blue-600">{obj.avg}</span>，整体达成情况评价为“<span className="font-bold text-blue-700">{obj.conclusion}</span>”。
                            从分布图来看，大部分学生的达成评价值集中在 {Math.max(0, obj.avg - 0.15).toFixed(2)} 至 {Math.min(1, obj.avg + 0.15).toFixed(2)} 区间内，反映出教学过程对该目标的支撑作用较为稳固。
                            针对现状分析：{obj.lowCount > 0 
                              ? `目前仍有 ${obj.lowCount} 名学生未能达到合格标准，主要原因可能在于对“${courseInfo.objectiveDescriptions[i].substring(0, 25)}...”相关知识点的理解深度不足或应用能力欠缺。在后续教学中，需针对这部分学生加强个别辅导，并优化相关教学环节的互动性。` 
                              : `全体学生均已达到合格及以上标准，其中优秀率达到 ${Math.round((obj.highCount / finalData.length) * 100)}%，说明教学方法与考核方式高度契合，学生对该目标的掌握情况非常理想。`}
                            总结而言，该目标达成的现状符合预期教学要求，但仍需关注个体差异，持续优化教学质量。
                          </p>
                        </div>
                      </div>
                    </div>
                    {!isExporting && (
                      <button
                        onClick={() => downloadChart(`chart-objective-${i}`, `课程目标${obj.index}达成情况分布图`)}
                        className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 text-slate-600 rounded-lg hover:bg-slate-200 transition-all text-xs font-bold shrink-0"
                      >
                        <Download className="w-3.5 h-3.5" />
                        下载图片
                      </button>
                    )}
                  </div>
                  
                  <div id={`chart-objective-${i}`} className="h-[350px] w-full bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <ResponsiveContainer width="100%" height="100%">
                        <ScatterChart 
                          margin={isExporting ? { top: 20, right: 100, bottom: 2, left: 70 } : { top: 20, right: 30, bottom: 20, left: 20 }}
                        >
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis 
                          type="number" 
                          dataKey="x" 
                          name="学生序号" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ 
                            fontSize: isExporting ? 20 : 10,
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                        <YAxis 
                          type="number" 
                          dataKey="y" 
                          name="达成评价值" 
                          domain={[0, 1]} 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ 
                            fontSize: isExporting ? 20 : 10,
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                        <ZAxis type="number" range={[30, 30]} />
                        <Tooltip 
                          cursor={{ strokeDasharray: '3 3' }} 
                          contentStyle={{ fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif" }}
                        />
                        <Scatter name={`目标${obj.index}达成度`} data={obj.data} fill={COLORS[i % COLORS.length]} isAnimationActive={!isExporting} />
                        <ReferenceLine 
                          y={obj.avg} 
                          stroke="#3b82f6" 
                          strokeWidth={2} 
                          label={{ 
                            position: 'right', 
                            value: `平均值: ${obj.avg}`, 
                            fill: '#3b82f6', 
                            fontSize: isExporting ? 24 : 12, 
                            fontWeight: 'bold',
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                        <ReferenceLine 
                          y={0.65} 
                          stroke="#ef4444" 
                          strokeDasharray="5 5" 
                          label={{ 
                            position: 'left', 
                            value: '0.65 警戒线', 
                            fill: '#ef4444', 
                            fontSize: isExporting ? 20 : 10,
                            fontFamily: "'Times New Roman', 'SimSun', 'SimHei', sans-serif"
                          }} 
                        />
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>
                  <p className="text-center text-xs text-slate-400 mt-4">图{obj.index} 课程目标{obj.index}学生达成情况分布图</p>
                </motion.section>
              ))}
            </div>
          )}
        </AnimatePresence>

        {isExporting && (
          <section className="bg-white p-0 rounded-none border-none shadow-none break-before-page">
            <SectionHeader icon={FileText} title={isExporting ? "六、课程目标达成情况评价与课程质量持续改进措施" : "七、课程目标达成情况评价与课程质量持续改进措施"} step={isExporting ? 6 : 7} />
            <div className="overflow-x-auto">
              <table className="w-full border-collapse border border-black text-xs" style={{ border: '1px solid black' }}>
                <tbody>
                  <tr>
                    <td className="border border-black p-4 bg-slate-50 font-bold w-32 text-center" style={{ border: '1px solid black', width: '120px' }}>课程目标达成情况评价</td>
                    <td colSpan={3} className="border border-black p-4 leading-relaxed" style={{ border: '1px solid black', textAlign: 'left', textIndent: '2em', whiteSpace: 'pre-wrap', minHeight: '150px', verticalAlign: 'top' }}>
                      {evaluationMeasures.evaluation}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-black p-4 bg-slate-50 font-bold w-32 text-center" style={{ border: '1px solid black', width: '120px' }}>持续改进措施</td>
                    <td colSpan={3} className="border border-black p-4 leading-relaxed" style={{ border: '1px solid black', textAlign: 'left', textIndent: '0', whiteSpace: 'pre-wrap', minHeight: '150px', verticalAlign: 'top' }}>
                      {evaluationMeasures.improvement}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-black p-4 bg-slate-50 font-bold w-32 text-center" style={{ border: '1px solid black', width: '120px' }}>课程责任教授意见</td>
                    <td className="border border-black p-4 text-left" style={{ border: '1px solid black', width: '30%' }}>
                      {evaluationMeasures.professorOpinion || '措施合理'}
                    </td>
                    <td className="border border-black p-4 bg-slate-50 font-bold text-center" style={{ border: '1px solid black', width: '100px' }}>评价时间</td>
                    <td className="border border-black p-4 text-center" style={{ border: '1px solid black' }}>
                      {evaluationMeasures.evaluationDate}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-black p-4 bg-slate-50 font-bold w-32 text-center" style={{ border: '1px solid black', width: '120px' }}>课程群（组）责任教授审核意见</td>
                    <td colSpan={3} className="border border-black p-4" style={{ border: '1px solid black', verticalAlign: 'top' }}>
                      <div className="flex flex-col justify-between min-h-[120px]">
                        <p className="leading-relaxed text-left" style={{ textIndent: '2em' }}>{evaluationMeasures.auditOpinion}</p>
                        <div className="flex justify-between items-end mt-8" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end' }}>
                          <span>课程群（组）责任教授（签字）：</span>
                          <span>日期：{evaluationMeasures.auditDate}</span>
                        </div>
                      </div>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>
        )}

        </div>

        {/* Footer */}
        <footer className="text-center py-12 text-slate-400 text-sm border-t border-slate-200">
          <p>© 2026 大学课程目标达成度分析系统 · 基于 Google AI Studio 构建</p>
          <p className="mt-2">支持一键生成符合教学评估要求的分析报告</p>
        </footer>

      </div>

      {/* Loading Overlay */}
      {isProcessing && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-white p-8 rounded-2xl shadow-2xl flex flex-col items-center gap-4">
            <div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin"></div>
            <p className="font-bold text-slate-800">{loadingMessage}</p>
          </div>
        </div>
      )}

      {/* Visual Guide Overlay */}
      <AnimatePresence>
        {showGuide && (
          <div className="fixed inset-0 z-[100] pointer-events-none">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-black/40 pointer-events-auto"
              onClick={() => setShowGuide(false)}
            />
            
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="absolute z-[101] pointer-events-auto bg-white rounded-2xl shadow-2xl p-6 w-full max-w-sm border border-slate-200"
              style={{
                left: '50%',
                top: '50%',
                transform: 'translate(-50%, -50%)'
              }}
            >
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-full bg-blue-600 text-white flex items-center justify-center font-bold text-sm">
                    {guideStep + 1}
                  </div>
                  <h3 className="font-bold text-lg text-slate-900">{guideSteps[guideStep].title}</h3>
                </div>
                <button 
                  onClick={() => setShowGuide(false)}
                  className="p-1 hover:bg-slate-100 rounded-lg transition-colors"
                >
                  <X className="w-5 h-5 text-slate-400" />
                </button>
              </div>
              
              <p className="text-slate-600 leading-relaxed mb-6">
                {guideSteps[guideStep].content}
              </p>
              
              <div className="flex items-center justify-between">
                <div className="flex gap-1">
                  {guideSteps.map((_, i) => (
                    <div 
                      key={i} 
                      className={cn(
                        "w-2 h-2 rounded-full transition-all",
                        i === guideStep ? "bg-blue-600 w-4" : "bg-slate-200"
                      )}
                    />
                  ))}
                </div>
                <div className="flex gap-2">
                  {guideStep > 0 && (
                    <button 
                      onClick={prevGuideStep}
                      className="px-4 py-2 text-slate-600 font-semibold hover:bg-slate-50 rounded-xl transition-all"
                    >
                      上一步
                    </button>
                  )}
                  <button 
                    onClick={nextGuideStep}
                    className="px-6 py-2 bg-blue-600 text-white font-semibold rounded-xl hover:bg-blue-700 transition-all shadow-lg shadow-blue-200"
                  >
                    {guideStep === guideSteps.length - 1 ? "完成" : "下一步"}
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Validation Error Modal */}
      <AnimatePresence>
        {validationError && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              onClick={() => setValidationError(null)}
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative z-[111] bg-white rounded-2xl shadow-2xl p-8 w-full max-w-md border border-red-100"
            >
              <div className="flex items-center gap-4 text-red-600 mb-6">
                <div className="w-12 h-12 rounded-full bg-red-50 flex items-center justify-center">
                  <AlertCircle className="w-6 h-6" />
                </div>
                <h3 className="text-xl font-bold">导出校验未通过</h3>
              </div>
              
              <div className="p-4 bg-red-50 rounded-xl border border-red-100 mb-8 max-h-[300px] overflow-y-auto">
                <ul className="space-y-2">
                  {validationError.map((err, idx) => (
                    <li key={idx} className="text-red-800 leading-relaxed font-medium flex gap-2">
                      <span className="shrink-0">•</span>
                      <span>{err}</span>
                    </li>
                  ))}
                </ul>
              </div>
              
              <button 
                onClick={() => setValidationError(null)}
                className="w-full py-3 bg-slate-900 text-white font-bold rounded-xl hover:bg-slate-800 transition-all shadow-lg shadow-slate-200"
              >
                返回修改
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Confirmation Modal */}
      <AnimatePresence>
        {showConfirm && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
              onClick={() => setShowConfirm(null)}
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative z-[111] bg-white rounded-2xl shadow-2xl p-8 w-full max-w-md border border-blue-100"
            >
              <div className="flex items-center gap-4 text-blue-600 mb-6">
                <div className="w-12 h-12 rounded-full bg-blue-50 flex items-center justify-center">
                  <HelpCircle className="w-6 h-6" />
                </div>
                <h3 className="text-xl font-bold">导出确认</h3>
              </div>
              
              <div className="p-4 bg-blue-50 rounded-xl border border-blue-100 mb-8">
                <p className="text-blue-800 leading-relaxed font-medium">
                  {showConfirm.message}
                </p>
              </div>
              
              <div className="flex gap-3">
                <button 
                  onClick={() => setShowConfirm(null)}
                  className="flex-1 py-3 bg-slate-100 text-slate-600 font-bold rounded-xl hover:bg-slate-200 transition-all"
                >
                  取消
                </button>
                <button 
                  onClick={showConfirm.onConfirm}
                  className="flex-1 py-3 bg-blue-600 text-white font-bold rounded-xl hover:bg-blue-700 transition-all shadow-lg shadow-blue-200"
                >
                  确认继续
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
