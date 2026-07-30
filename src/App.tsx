import React, { useState, useCallback } from 'react'
import { useKV } from '@github/spark/hooks'
import { read as readXLSX, utils as xlsxUtils } from 'xlsx'
import { Upload, BarChart3, Download, FileSpreadsheet, TrendingUp, PieChart, Calendar, Clock, Funnel, Check, X, Filter, CalendarBlank, TrendDown, TrendUp , Wrench } from '@phosphor-icons/react'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Progress } from '@/components/ui/progress'
import { Badge } from '@/components/ui/badge'
import { Separator } from '@/components/ui/separator'
import { Checkbox } from '@/components/ui/checkbox'
import { Label } from '@/components/ui/label'
import { Input } from '@/components/ui/input'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Slider } from '@/components/ui/slider'
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from '@/components/ui/collapsible'
import { toast } from 'sonner'
import { DataTable } from '@/components/DataTable'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import * as ss from 'simple-statistics'
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, ResponsiveContainer, PieChart as RechartsPieChart, Cell, LineChart, Line, Pie, Tooltip, Legend, AreaChart, Area , ScatterChart, Scatter } from 'recharts'


const CustomTooltip = ({ active, payload, label }: any) => {
  if (active && payload && payload.length) {
    return (
      <div className="custom-chart-tooltip">
        <p className="font-semibold text-foreground">{label}</p>
        <p className="text-primary mt-1">
          {payload[0].name}: {payload[0].value}
        </p>
      </div>
    );
  }
  return null;
};

interface DataColumn {
  name: string
  type: 'number' | 'text' | 'date'
  values: any[]
  dateValues?: (Date | null)[] // Parsed dates for date columns - null for unparseable values
  detectedFormats?: string[] // Formats detected in date columns
  stats?: {
    min?: number | Date
    max?: number | Date
    avg?: number
    sum?: number
    stdDev?: number
    variance?: number
    count: number
    dateRange?: string // For date columns
    frequency?: 'daily' | 'weekly' | 'monthly' | 'yearly' | 'irregular'
    formatSummary?: string // Summary of detected date formats
  }
  // Track which rows are included in analysis
  selectedRows?: boolean[]
  originalRowCount?: number
}

interface Insight {
  type: 'trend' | 'correlation' | 'summary' | 'outlier' | 'temporal' | 'seasonal'
  title: string
  description: string
  value?: string | number
  importance: 'high' | 'medium' | 'low'
}

interface TimeSeriesData {
  date: string
  value: number
  originalDate: Date
}

interface FilterConfig {
  dateRange?: {
    start: string
    end: string
    columnIndex: number
  }
  outlierDetection?: {
    columnIndex: number
    sensitivity: number // 1-3 (1=conservative, 3=aggressive)
    method: 'iqr' | 'zscore' | 'modified_zscore'
    selectOutliers: boolean // true=select outliers, false=exclude outliers
  }
  valueRange?: {
    columnIndex: number
    min: number
    max: number
  }
  textFilter?: {
    columnIndex: number
    pattern: string
    caseSensitive: boolean
  }
}

const CHART_COLORS = ['#1e40af', '#f97316', '#10b981', '#8b5cf6', '#ef4444', '#06b6d4']

function App() {
  const [uploadedData, setUploadedData] = useKV<DataColumn[]>('excel-data', [])
  const [insights, setInsights] = useKV<Insight[]>('data-insights', [])
  const [isProcessing, setIsProcessing] = useState(false)
  const [processingStep, setProcessingStep] = useState('')

  const [customChartConfig, setCustomChartConfig] = useState<{
    xAxis: string;
    yAxis: string;
    chartType: 'bar' | 'line' | 'scatter' | 'area';
  }>({ xAxis: '', yAxis: '', chartType: 'bar' })
  const [fileName, setFileName] = useKV<string>('file-name', '')
  const [showDataSelection, setShowDataSelection] = useState(false)
  const [hasError, setHasError] = useState(false)
  const [filterConfig, setFilterConfig] = useKV<FilterConfig>('filter-config', {})
  const [showAdvancedFilters, setShowAdvancedFilters] = useState(false)

  // Safe data access with fallbacks
  const safeUploadedData = Array.isArray(uploadedData) ? uploadedData : []
  const safeInsights = Array.isArray(insights) ? insights : []
  const safeFilterConfig = filterConfig || {}

  // Error boundary for component
  if (hasError) {
    return (
      <div className="min-h-screen bg-background p-6 flex items-center justify-center">
        <Card className="max-w-md mx-auto">
          <CardHeader>
            <CardTitle className="text-center text-destructive">Application Error</CardTitle>
          </CardHeader>
          <CardContent className="text-center space-y-4">
            <p className="text-muted-foreground">
              An error occurred while processing your data. Please try refreshing the page or uploading your file again.
            </p>
            <Button onClick={() => {
              setHasError(false)
              setUploadedData([])
              setInsights([])
              setFileName('')
              setFilterConfig({})
            }}>
              Reset Application
            </Button>
          </CardContent>
        </Card>
      </div>
    )
  }

  // Enhanced helper function to detect and parse various date formats
  const parseDate = useCallback((value: any): Date | null => {
    if (!value) return null
    
    // Handle Excel serial dates
    if (typeof value === 'number' && value > 1 && value < 100000) {
      // Excel serial date (days since 1900-01-01, with leap year bug)
      const excelEpoch = new Date(1900, 0, 1)
      const date = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000)
      return isNaN(date.getTime()) ? null : date
    }
    
    // Handle string dates
    if (typeof value === 'string') {
      const cleanValue = value.trim()
      
      // Comprehensive date format patterns with explicit parsing
      const dateFormats = [
        // ISO and standard formats
        { pattern: /^\d{4}-\d{2}-\d{2}$/, format: 'YYYY-MM-DD' },
        { pattern: /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/, format: 'ISO' },
        { pattern: /^\d{4}\/\d{2}\/\d{2}$/, format: 'YYYY/MM/DD' },
        
        // US formats (MM/DD/YYYY variations)
        { pattern: /^\d{1,2}\/\d{1,2}\/\d{4}$/, format: 'M/D/YYYY' },
        { pattern: /^\d{2}\/\d{2}\/\d{4}$/, format: 'MM/DD/YYYY' },
        { pattern: /^\d{1,2}-\d{1,2}-\d{4}$/, format: 'M-D-YYYY' },
        { pattern: /^\d{2}-\d{2}-\d{4}$/, format: 'MM-DD-YYYY' },
        
        // European formats (DD/MM/YYYY variations)
        { pattern: /^\d{1,2}\/\d{1,2}\/\d{4}$/, format: 'D/M/YYYY' },
        { pattern: /^\d{2}\.\d{2}\.\d{4}$/, format: 'DD.MM.YYYY' },
        { pattern: /^\d{1,2}\.\d{1,2}\.\d{4}$/, format: 'D.M.YYYY' },
        
        // Alternative separators
        { pattern: /^\d{4}\.\d{2}\.\d{2}$/, format: 'YYYY.MM.DD' },
        { pattern: /^\d{2}\s+\d{2}\s+\d{4}$/, format: 'MM DD YYYY' },
        
        // With time components
        { pattern: /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}/, format: 'M/D/YYYY H:MM' },
        { pattern: /^\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}/, format: 'YYYY-MM-DD H:MM' }
      ]
      
      // Try each format pattern
      for (const { pattern, format } of dateFormats) {
        if (pattern.test(cleanValue)) {
          let parsedDate: Date | null = null
          
          // Custom parsing for ambiguous formats
          if (format === 'M/D/YYYY' || format === 'MM/DD/YYYY') {
            // US format: Month/Day/Year
            const parts = cleanValue.split('/')
            if (parts.length === 3) {
              const month = parseInt(parts[0], 10) - 1 // 0-based month
              const day = parseInt(parts[1], 10)
              const year = parseInt(parts[2], 10)
              
              // Validate month and day ranges for US format
              if (month >= 0 && month <= 11 && day >= 1 && day <= 31) {
                parsedDate = new Date(year, month, day)
              }
            }
          } else if (format === 'D/M/YYYY') {
            // European format: Day/Month/Year (for ambiguous cases)
            const parts = cleanValue.split('/')
            if (parts.length === 3) {
              const day = parseInt(parts[0], 10)
              const month = parseInt(parts[1], 10) - 1 // 0-based month
              const year = parseInt(parts[2], 10)
              
              // Validate day and month ranges for European format
              if (day >= 1 && day <= 31 && month >= 0 && month <= 11) {
                parsedDate = new Date(year, month, day)
              }
            }
          } else if (format === 'DD.MM.YYYY' || format === 'D.M.YYYY') {
            // European format: Day.Month.Year
            const parts = cleanValue.split('.')
            if (parts.length === 3) {
              const day = parseInt(parts[0], 10)
              const month = parseInt(parts[1], 10) - 1 // 0-based month
              const year = parseInt(parts[2], 10)
              parsedDate = new Date(year, month, day)
            }
          } else {
            // Let JavaScript's Date constructor handle standard formats
            parsedDate = new Date(cleanValue)
          }
          
          // Validate the parsed date
          if (parsedDate && !isNaN(parsedDate.getTime())) {
            // Additional validation: ensure reasonable date range
            const year = parsedDate.getFullYear()
            if (year >= 1900 && year <= 2100) {
              return parsedDate
            }
          }
        }
      }
      
      // Try month names (e.g., "January 15, 2023", "15 Jan 2023")
      const monthNamePatterns = [
        /^\w+\s+\d{1,2},?\s+\d{4}$/i, // "January 15, 2023" or "January 15 2023"
        /^\d{1,2}\s+\w+\s+\d{4}$/i,   // "15 January 2023"
        /^\w+\s+\d{4}$/i,             // "January 2023"
      ]
      
      for (const pattern of monthNamePatterns) {
        if (pattern.test(cleanValue)) {
          const date = new Date(cleanValue)
          if (!isNaN(date.getTime())) {
            return date
          }
        }
      }
      
      // Final fallback: try Date.parse for any remaining formats
      const timestamp = Date.parse(cleanValue)
      if (!isNaN(timestamp)) {
        const date = new Date(timestamp)
        const year = date.getFullYear()
        if (year >= 1900 && year <= 2100) {
          return date
        }
      }
    }
    
    // Handle Date objects
    if (value instanceof Date) {
      return isNaN(value.getTime()) ? null : value
    }
    
    return null
  }, [])

  // Analyze date frequency and patterns
  const analyzeDateFrequency = useCallback((dates: Date[]): 'daily' | 'weekly' | 'monthly' | 'yearly' | 'irregular' => {
    if (dates.length < 2) return 'irregular'
    
    const sortedDates = [...dates].sort((a, b) => a.getTime() - b.getTime())
    const intervals: number[] = []
    
    for (let i = 1; i < sortedDates.length; i++) {
      const diff = sortedDates[i].getTime() - sortedDates[i - 1].getTime()
      intervals.push(diff)
    }
    
    const avgInterval = intervals.reduce((a, b) => a + b, 0) / intervals.length
    const dayMs = 24 * 60 * 60 * 1000
    
    if (avgInterval <= dayMs * 1.5) return 'daily'
    if (avgInterval <= dayMs * 8) return 'weekly'
    if (avgInterval <= dayMs * 35) return 'monthly'
    if (avgInterval <= dayMs * 400) return 'yearly'
    
    return 'irregular'
  }, [])

  // Calculate statistics for selected data points only
  const calculateStats = useCallback((column: DataColumn): DataColumn['stats'] => {
    // Add null check for column parameter
    if (!column || !column.selectedRows || !column.values) {
      return { count: 0 }
    }
    
    const selectedIndices = column.selectedRows
      .map((selected, index) => selected ? index : -1)
      .filter(index => index !== -1)
    
    if (selectedIndices.length === 0) {
      return { count: 0 }
    }

    if (column.type === 'date' && column.dateValues) {
      const selectedDates = selectedIndices
        .map(index => column.dateValues![index])
        .filter((date): date is Date => date !== null)
      
      if (selectedDates.length === 0) return { count: 0 }
      
      const sortedDates = [...selectedDates].sort((a, b) => a.getTime() - b.getTime())
      const frequency = analyzeDateFrequency(selectedDates)
      
      const minDate = sortedDates[0]
      const maxDate = sortedDates[sortedDates.length - 1]
      const dateRange = minDate && maxDate && 
        !isNaN(minDate.getTime()) && !isNaN(maxDate.getTime()) ? 
        `${minDate.toLocaleDateString()} to ${maxDate.toLocaleDateString()}` : 
        'Invalid range'
      
      return {
        count: selectedDates.length,
        min: minDate,
        max: maxDate,
        dateRange,
        frequency,
        formatSummary: column.stats?.formatSummary // Keep original format summary
      }
    }

    if (column.type === 'number') {
      const selectedValues = selectedIndices
        .map(index => Number(column.values[index]))
        .filter(val => !isNaN(val) && isFinite(val))
      
      if (selectedValues.length === 0) return { count: 0 }
      
      return {
        count: selectedValues.length,
        min: Math.min(...selectedValues),
        max: Math.max(...selectedValues),
        avg: selectedValues.reduce((a, b) => a + b, 0) / selectedValues.length,
        sum: selectedValues.reduce((a, b) => a + b, 0),
        stdDev: selectedValues.length > 1 ? ss.standardDeviation(selectedValues) : 0,
        variance: selectedValues.length > 1 ? ss.variance(selectedValues) : 0
      }
    }

    return { count: selectedIndices.length }
  }, [analyzeDateFrequency])

  // Advanced outlier detection methods
  const detectOutliers = useCallback((values: number[], method: 'iqr' | 'zscore' | 'modified_zscore', sensitivity: number): boolean[] => {
    if (values.length < 4) return new Array(values.length).fill(false)
    
    const sortedValues = [...values].sort((a, b) => a - b)
    const outliers = new Array(values.length).fill(false)
    
    if (method === 'iqr') {
      const q1Index = Math.floor(sortedValues.length * 0.25)
      const q3Index = Math.floor(sortedValues.length * 0.75)
      const q1 = sortedValues[q1Index]
      const q3 = sortedValues[q3Index]
      const iqr = q3 - q1
      
      // Sensitivity: 1=1.5*IQR, 2=1.25*IQR, 3=1.0*IQR
      const multiplier = 2 - (sensitivity - 1) * 0.25
      const lowerBound = q1 - multiplier * iqr
      const upperBound = q3 + multiplier * iqr
      
      values.forEach((value, index) => {
        outliers[index] = value < lowerBound || value > upperBound
      })
    } else if (method === 'zscore') {
      const mean = values.reduce((sum, val) => sum + val, 0) / values.length
      const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length
      const stdDev = Math.sqrt(variance)
      
      // Sensitivity: 1=3σ, 2=2.5σ, 3=2σ
      const threshold = 4 - sensitivity
      
      values.forEach((value, index) => {
        const zScore = Math.abs((value - mean) / stdDev)
        outliers[index] = zScore > threshold
      })
    } else if (method === 'modified_zscore') {
      const median = sortedValues[Math.floor(sortedValues.length / 2)]
      const medianAbsoluteDeviations = values.map(val => Math.abs(val - median))
      const mad = medianAbsoluteDeviations.sort((a, b) => a - b)[Math.floor(medianAbsoluteDeviations.length / 2)]
      
      // Sensitivity: 1=3.5, 2=3.0, 3=2.5
      const threshold = 4 - (sensitivity - 1) * 0.5
      
      values.forEach((value, index) => {
        const modifiedZScore = 0.6745 * (value - median) / mad
        outliers[index] = Math.abs(modifiedZScore) > threshold
      })
    }
    
    return outliers
  }, [])

  // Apply date range filter
  const applyDateRangeFilter = useCallback((data: DataColumn[], config: FilterConfig['dateRange']): DataColumn[] => {
    if (!config || !data[config.columnIndex] || data[config.columnIndex].type !== 'date') {
      return data
    }
    
    const startDate = new Date(config.start)
    const endDate = new Date(config.end)
    const column = data[config.columnIndex]
    
    if (!column.dateValues || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return data
    }
    
    const updatedData = [...data]
    const updatedColumn = { ...column }
    
    updatedColumn.selectedRows = column.dateValues.map((date, index) => {
      if (!date || !(date instanceof Date) || isNaN(date.getTime())) return false
      return date >= startDate && date <= endDate && (column.selectedRows?.[index] ?? true)
    })
    
    updatedData[config.columnIndex] = updatedColumn
    return updatedData
  }, [])

  // Apply outlier detection filter
  const applyOutlierFilter = useCallback((data: DataColumn[], config: FilterConfig['outlierDetection']): DataColumn[] => {
    if (!config || !data[config.columnIndex] || data[config.columnIndex].type !== 'number') {
      return data
    }
    
    const column = data[config.columnIndex]
    const numericValues = column.values
      .map(val => Number(val))
      .filter(val => !isNaN(val) && isFinite(val))
    
    if (numericValues.length < 4) return data
    
    const outlierFlags = detectOutliers(numericValues, config.method, config.sensitivity)
    
    const updatedData = [...data]
    const updatedColumn = { ...column }
    
    let outlierIndex = 0
    updatedColumn.selectedRows = column.values.map((val, index) => {
      const numVal = Number(val)
      if (isNaN(numVal) || !isFinite(numVal)) return column.selectedRows?.[index] ?? true
      
      const isOutlier = outlierFlags[outlierIndex++]
      // Select outliers if selectOutliers is true, otherwise exclude them
      const shouldSelect = config.selectOutliers ? isOutlier : !isOutlier
      return shouldSelect && (column.selectedRows?.[index] ?? true)
    })
    
    updatedData[config.columnIndex] = updatedColumn
    return updatedData
  }, [detectOutliers])

  // Apply value range filter
  const applyValueRangeFilter = useCallback((data: DataColumn[], config: FilterConfig['valueRange']): DataColumn[] => {
    if (!config || !data[config.columnIndex] || data[config.columnIndex].type !== 'number') {
      return data
    }
    
    const column = data[config.columnIndex]
    const updatedData = [...data]
    const updatedColumn = { ...column }
    
    updatedColumn.selectedRows = column.values.map((val, index) => {
      const numVal = Number(val)
      if (isNaN(numVal) || !isFinite(numVal)) return false
      return numVal >= config.min && numVal <= config.max && (column.selectedRows?.[index] ?? true)
    })
    
    updatedData[config.columnIndex] = updatedColumn
    return updatedData
  }, [])

  // Apply text pattern filter
  const applyTextFilter = useCallback((data: DataColumn[], config: FilterConfig['textFilter']): DataColumn[] => {
    if (!config || !data[config.columnIndex] || !config.pattern) {
      return data
    }
    
    const column = data[config.columnIndex]
    const updatedData = [...data]
    const updatedColumn = { ...column }
    
    const pattern = config.caseSensitive ? config.pattern : config.pattern.toLowerCase()
    
    updatedColumn.selectedRows = column.values.map((val, index) => {
      if (val === null || val === undefined) return false
      const strVal = config.caseSensitive ? String(val) : String(val).toLowerCase()
      return strVal.includes(pattern) && (column.selectedRows?.[index] ?? true)
    })
    
    updatedData[config.columnIndex] = updatedColumn
    return updatedData
  }, [])

  // Apply all active filters
  const applyAdvancedFilters = useCallback((baseData: DataColumn[]): DataColumn[] => {
    let filteredData = [...baseData]
    
    if (safeFilterConfig.dateRange) {
      filteredData = applyDateRangeFilter(filteredData, safeFilterConfig.dateRange)
    }
    
    if (safeFilterConfig.outlierDetection) {
      filteredData = applyOutlierFilter(filteredData, safeFilterConfig.outlierDetection)
    }
    
    if (safeFilterConfig.valueRange) {
      filteredData = applyValueRangeFilter(filteredData, safeFilterConfig.valueRange)
    }
    
    if (safeFilterConfig.textFilter) {
      filteredData = applyTextFilter(filteredData, safeFilterConfig.textFilter)
    }
    
    return filteredData
  }, [safeFilterConfig, applyDateRangeFilter, applyOutlierFilter, applyValueRangeFilter, applyTextFilter])

  // Update filter configuration
  const updateFilterConfig = useCallback((newConfig: Partial<FilterConfig>) => {
    const updatedConfig = { ...safeFilterConfig, ...newConfig }
    setFilterConfig(updatedConfig)
    
    if (safeUploadedData.length > 0) {
      const filteredData = applyAdvancedFilters(safeUploadedData)
      const columnsWithUpdatedStats = filteredData.map(col => ({
        ...col,
        stats: calculateStats(col)
      }))
      
      const newInsights = analyzeData(columnsWithUpdatedStats)
      setUploadedData(columnsWithUpdatedStats)
      setInsights(newInsights)
      
      toast.success('Filters applied successfully')
    }
  }, [safeFilterConfig, safeUploadedData, applyAdvancedFilters, calculateStats, setFilterConfig, setUploadedData, setInsights])

  // Generate time-series insights
  const generateTimeSeriesInsights = useCallback((dateColumns: DataColumn[], numericColumns: DataColumn[]): Insight[] => {
    const insights: Insight[] = []
    
    // Find time series pairs (date + numeric columns)
    dateColumns.forEach(dateCol => {
      if (!dateCol.dateValues || dateCol.dateValues.length === 0) return
      
      const dateRange = dateCol.stats?.dateRange
      const frequency = dateCol.stats?.frequency
      
      // Date range insight with format information
      insights.push({
        type: 'temporal',
        title: `${dateCol.name} Date Analysis`,
        description: `Data spans ${dateRange} with ${frequency} frequency. Detected formats: ${dateCol.stats?.formatSummary || 'Various'}`,
        value: dateRange,
        importance: 'high'
      })
      
      // Frequency pattern insight
      if (frequency !== 'irregular') {
        insights.push({
          type: 'temporal',
          title: 'Data Collection Pattern',
          description: `Regular ${frequency} data collection detected, ideal for trend analysis and forecasting`,
          importance: 'medium'
        })
      }
      
      // Seasonal analysis for monthly/yearly data
      if (dateCol.dateValues && dateCol.dateValues.filter(d => d !== null).length >= 12 && (frequency === 'monthly' || frequency === 'yearly')) {
        const monthCounts = new Array(12).fill(0)
        dateCol.dateValues.forEach(date => {
          if (date) {
            monthCounts[date.getMonth()]++
          }
        })
        
        const maxMonth = monthCounts.indexOf(Math.max(...monthCounts))
        const minMonth = monthCounts.indexOf(Math.min(...monthCounts))
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        
        insights.push({
          type: 'seasonal',
          title: 'Seasonal Patterns',
          description: `Peak activity in ${monthNames[maxMonth]}, lowest in ${monthNames[minMonth]}`,
          importance: 'medium'
        })
      }
      
      // Combine with numeric data for trend analysis
      numericColumns.forEach(numCol => {
        if (dateCol.values.length === numCol.values.length) {
          // Create time series pairs
          const pairs: Array<{date: Date, value: number}> = []
          for (let i = 0; i < dateCol.dateValues!.length; i++) {
            const date = dateCol.dateValues![i]
            const value = Number(numCol.values[i])
            if (date && !isNaN(value)) {
              pairs.push({ date, value })
            }
          }
          
          if (pairs.length >= 3) {
            // Sort by date
            pairs.sort((a, b) => a.date.getTime() - b.date.getTime())
            
            // Calculate trend
            const firstHalf = pairs.slice(0, Math.floor(pairs.length / 2))
            const secondHalf = pairs.slice(Math.floor(pairs.length / 2))
            
            const firstAvg = firstHalf.reduce((sum, p) => sum + p.value, 0) / firstHalf.length
            const secondAvg = secondHalf.reduce((sum, p) => sum + p.value, 0) / secondHalf.length
            
            const trendDirection = secondAvg > firstAvg ? 'increasing' : 'decreasing'
            const trendMagnitude = Math.abs((secondAvg - firstAvg) / firstAvg * 100)
            
            if (trendMagnitude > 5) {
              insights.push({
                type: 'trend',
                title: `${numCol.name} Time Trend`,
                description: `${numCol.name} shows ${trendDirection} trend over time with ${trendMagnitude.toFixed(1)}% change`,
                value: `${trendDirection === 'increasing' ? '+' : '-'}${trendMagnitude.toFixed(1)}%`,
                importance: trendMagnitude > 20 ? 'high' : 'medium'
              })
            }
          }
        }
      })
    })
    
    return insights
  }, [])

  const analyzeData = useCallback((columns: DataColumn[]): Insight[] => {
    const newInsights: Insight[] = []
    
    // Filter out null/undefined columns and ensure they have required properties
    const validColumns = columns.filter(col => 
      col && 
      col.values && 
      col.selectedRows && 
      col.name && 
      col.type
    )
    
    // Filter columns to only include those with selected data
    const activeColumns = validColumns.filter(col => {
      const selectedCount = col.selectedRows?.filter(Boolean).length || 0
      return selectedCount > 0
    })
    
    if (activeColumns.length === 0) {
      newInsights.push({
        type: 'summary',
        title: 'No Data Selected',
        description: 'Please select data points to generate insights and statistics.',
        importance: 'high'
      })
      return newInsights
    }
    
    // Find different column types among active columns
    const numericColumns = activeColumns.filter(col => col.type === 'number' && col.stats)
    const dateColumns = activeColumns.filter(col => col.type === 'date')
    const textColumns = activeColumns.filter(col => col.type === 'text')
    
    // Calculate total selected rows
    const totalSelectedRows = Math.max(...activeColumns.map(col => 
      col.selectedRows?.filter(Boolean).length || 0
    ))
    
    // Add selection overview insight
    const totalOriginalRows = activeColumns[0]?.originalRowCount || 0
    const selectionPercentage = totalOriginalRows > 0 ? 
      Math.round((totalSelectedRows / totalOriginalRows) * 100) : 0
    
    newInsights.push({
      type: 'summary',
      title: 'Data Selection Overview',
      description: `Analyzing ${totalSelectedRows} selected rows (${selectionPercentage}% of original data) across ${activeColumns.length} columns: ${numericColumns.length} numeric, ${dateColumns.length} date, ${textColumns.length} text`,
      value: `${totalSelectedRows}/${totalOriginalRows} records`,
      importance: 'high'
    })
    
    // Generate time-series insights if date columns exist
    if (dateColumns.length > 0) {
      const timeSeriesInsights = generateTimeSeriesInsights(dateColumns, numericColumns)
      newInsights.push(...timeSeriesInsights)
      
      // Add date format detection summary for selected data
      const allFormats = dateColumns
        .flatMap(col => col.detectedFormats || [])
        .reduce((acc, format) => {
          acc[format] = (acc[format] || 0) + 1
          return acc
        }, {} as Record<string, number>)
      
      const uniqueFormats = Object.keys(allFormats)
      const formatsList = uniqueFormats.join(', ')
      
      if (formatsList && uniqueFormats.length > 1) {
        newInsights.push({
          type: 'summary',
          title: 'Selected Data Date Formats',
          description: `Selected data contains ${uniqueFormats.length} different date formats: ${formatsList}. All dates normalized for consistent analysis.`,
          value: `${uniqueFormats.length} formats`,
          importance: 'medium'
        })
      }
    }
    
    if (numericColumns.length > 0) {
      // Find highest value column among selected data
      const maxColumn = numericColumns.reduce((max, col) => 
        (col.stats!.max as number || 0) > (max.stats!.max as number || 0) ? col : max
      )
      
      newInsights.push({
        type: 'trend',
        title: 'Highest Selected Values',
        description: `${maxColumn.name} shows the highest peak value in your selected dataset`,
        value: (maxColumn.stats!.max as number)?.toLocaleString(),
        importance: 'high'
      })

      // Find average insights for selected data
      numericColumns.forEach(col => {
        if (col.stats!.avg && col.stats!.avg > 0) {
          newInsights.push({
            type: 'summary',
            title: `${col.name} Average (Selected)`,
            description: `The average ${col.name.toLowerCase()} across selected records`,
            value: col.stats!.avg.toLocaleString(undefined, { maximumFractionDigits: 2 }),
            importance: 'medium'
          })
        }
      })

      // Distribution insight for selected data
      if (numericColumns.length >= 2) {
        newInsights.push({
          type: 'correlation',
          title: 'Selected Data Distribution',
          description: `Your selected data spans multiple metrics, enabling comprehensive analysis across ${numericColumns.length} dimensions`,
          importance: 'medium'
        })
      }
    } else if (textColumns.length > 0) {
      // Add insights for selected text-based data
      newInsights.push({
        type: 'summary',
        title: 'Selected Text Data',
        description: `Your selected dataset contains ${textColumns.length} text columns. Consider formatting numeric data as numbers for statistical analysis.`,
        importance: 'medium'
      })
    }

    return newInsights.slice(0, 8) // Limit to 8 insights
  }, [generateTimeSeriesInsights])

  // Clear all filters
  const clearAllFilters = useCallback(() => {
    setFilterConfig({})
    
    if (safeUploadedData.length > 0) {
      // Reset all rows to selected
      const resetData = safeUploadedData.map(col => ({
        ...col,
        selectedRows: new Array(col.values.length).fill(true)
      }))
      
      const columnsWithUpdatedStats = resetData.map(col => ({
        ...col,
        stats: calculateStats(col)
      }))
      
      const newInsights = analyzeData(columnsWithUpdatedStats)
      setUploadedData(columnsWithUpdatedStats)
      setInsights(newInsights)
      
      toast.success('All filters cleared')
    }
  }, [safeUploadedData, calculateStats, analyzeData, setFilterConfig, setUploadedData, setInsights])


  const processExcelFile = useCallback(async (file: File) => {
    setIsProcessing(true)
    setProcessingStep('Reading file...')
    
    try {
      const buffer = await file.arrayBuffer()
      setProcessingStep('Parsing Excel data...')
      
      const workbook = readXLSX(buffer, { type: 'buffer' })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]
      const jsonData = xlsxUtils.sheet_to_json(worksheet, { header: 1 })
      
      if (jsonData.length === 0) {
        throw new Error('No data found in the Excel file')
      }

      setProcessingStep('Analyzing data structure...')
      
      // Extract headers and data
      const headers = jsonData[0] as string[]
      const rows = jsonData.slice(1) as any[][]
      
      // Analyze each column
      const columns: DataColumn[] = headers.map((header, index) => {
        const values = rows.map(row => row[index]).filter(val => val !== undefined && val !== null && val !== '')
        
        // Try to detect dates first - maintain index alignment
        const dateValues: (Date | null)[] = []
        const detectedFormats: string[] = []
        let dateCount = 0
        
        values.forEach(val => {
          const parsedDate = parseDate(val)
          if (parsedDate) {
            dateValues.push(parsedDate)
            dateCount++
            
            // Track format for reporting
            if (typeof val === 'string') {
              const valStr = val.trim()
              if (/^\d{4}-\d{2}-\d{2}$/.test(valStr)) detectedFormats.push('YYYY-MM-DD')
              else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(valStr)) detectedFormats.push('MM/DD/YYYY')
              else if (/^\d{2}\.\d{2}\.\d{4}$/.test(valStr)) detectedFormats.push('DD.MM.YYYY')
              else if (/^\d{4}\/\d{2}\/\d{2}$/.test(valStr)) detectedFormats.push('YYYY/MM/DD')
              else if (/^\w+\s+\d{1,2},?\s+\d{4}$/i.test(valStr)) detectedFormats.push('Month DD, YYYY')
              else detectedFormats.push('Other')
            } else if (typeof val === 'number') {
              detectedFormats.push('Excel Serial')
            }
          } else {
            dateValues.push(null) // Maintain index alignment
          }
        })
        
        // Debug logging for date detection
        if (dateCount > 0) {
          console.log(`Date analysis for column "${header}":`, {
            totalValues: values.length,
            datesParsed: dateCount,
            successRate: `${((dateCount / values.length) * 100).toFixed(1)}%`,
            sampleValues: values.slice(0, 5),
            detectedFormats: [...new Set(detectedFormats)]
          })
        }
        
        // Determine if this is a date column (at least 60% are valid dates)
        const isDateColumn = values.length > 0 && (dateCount / values.length) >= 0.6
        
        if (isDateColumn) {
          // This is a date column
          const validDates = dateValues.filter((date): date is Date => date !== null)
          const sortedDates = [...validDates].sort((a, b) => a.getTime() - b.getTime())
          const frequency = analyzeDateFrequency(validDates)
          
          const minDate = sortedDates[0]
          const maxDate = sortedDates[sortedDates.length - 1]
          const dateRange = minDate && maxDate && 
            !isNaN(minDate.getTime()) && !isNaN(maxDate.getTime()) ? 
            `${minDate.toLocaleDateString()} to ${maxDate.toLocaleDateString()}` : 
            'Invalid range'
          
          // Create format summary
          const formatCounts = detectedFormats.reduce((acc, format) => {
            acc[format] = (acc[format] || 0) + 1
            return acc
          }, {} as Record<string, number>)
          
          const formatSummary = Object.entries(formatCounts)
            .sort(([,a], [,b]) => b - a)
            .map(([format, count]) => `${format} (${count})`)
            .join(', ')
          
          return {
            name: header,
            type: 'date' as const,
            values,
            dateValues,
            detectedFormats,
            stats: {
              count: values.length,
              min: minDate,
              max: maxDate,
              dateRange,
              frequency,
              formatSummary
            },
            selectedRows: new Array(values.length).fill(true), // Initially select all rows
            originalRowCount: values.length
          }
        }
        
        // Check if numeric (at least 80% of values are numeric)
        const numericCount = values.filter(val => {
          const num = Number(val)
          return !isNaN(num) && isFinite(num)
        }).length
        const isNumeric = values.length > 0 && (numericCount / values.length) >= 0.8
        const type = isNumeric ? 'number' : 'text'
        
        // Calculate statistics for numeric columns
        let stats: DataColumn['stats'] = { count: values.length }
        
        if (type === 'number') {
          const numericValues = values
            .map(val => Number(val))
            .filter(val => !isNaN(val) && isFinite(val))
          
          if (numericValues.length > 0) {
            stats = {
              count: values.length,
              min: Math.min(...numericValues),
              max: Math.max(...numericValues),
              avg: numericValues.reduce((a, b) => a + b, 0) / numericValues.length,
              sum: numericValues.reduce((a, b) => a + b, 0),
              stdDev: numericValues.length > 1 ? ss.standardDeviation(numericValues) : 0,
              variance: numericValues.length > 1 ? ss.variance(numericValues) : 0
            }
          }
        }
        
        return {
          name: header,
          type,
          values,
          stats,
          selectedRows: new Array(values.length).fill(true), // Initially select all rows
          originalRowCount: values.length
        }
      })

      setProcessingStep('Generating insights...')
      
      // Generate insights
      const generatedInsights = analyzeData(columns)
      
      console.log('Processed columns:', columns.map(col => ({ 
        name: col.name, 
        type: col.type, 
        count: col.stats?.count,
        detectedFormats: col.type === 'date' ? col.detectedFormats : undefined,
        formatSummary: col.type === 'date' ? col.stats?.formatSummary : undefined
      })))
      console.log('Generated insights:', generatedInsights)
      
      // Update state
      setUploadedData(columns)
      setInsights(generatedInsights)
      setFileName(file.name)
      
      toast.success(`File processed successfully! Found ${columns.length} columns and generated ${generatedInsights.length} insights.`)
      
    } catch (error) {
      toast.error(`Error processing file: ${error instanceof Error ? error.message : 'Unknown error'}`)
    } finally {
      setIsProcessing(false)
      setProcessingStep('')
    }
  }, [analyzeData, setUploadedData, setInsights, setFileName, parseDate, analyzeDateFrequency])

  const handleFileUpload = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (file) {
      if (file.size > 10 * 1024 * 1024) { // 10MB limit
        toast.error('File size must be less than 10MB')
        return
      }
      
      const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv'
      ]
      
      if (!validTypes.includes(file.type)) {
        toast.error('Please upload an Excel file (.xlsx, .xls) or CSV file')
        return
      }
      
      processExcelFile(file)
    }
  }, [processExcelFile])

  // Update data selection and recalculate insights
  const updateDataSelection = useCallback((updatedColumns: DataColumn[]) => {
    try {
      // Filter out any null/undefined columns and add safety checks
      const validColumns = updatedColumns.filter(col => col && col.values && col.selectedRows)
      
      if (validColumns.length === 0) {
        console.warn('No valid columns found for data selection update')
        return
      }
      
      // Recalculate stats for all columns based on selected rows
      const columnsWithUpdatedStats = validColumns.map(col => ({
        ...col,
        stats: calculateStats(col)
      }))
      
      // Regenerate insights based on selected data
      const newInsights = analyzeData(columnsWithUpdatedStats)
      
      setUploadedData(columnsWithUpdatedStats)
      setInsights(newInsights)
      
      const selectedCount = columnsWithUpdatedStats[0]?.selectedRows?.filter(Boolean).length || 0
      const totalCount = columnsWithUpdatedStats[0]?.originalRowCount || 0
      
      toast.success(`Analysis updated with ${selectedCount}/${totalCount} selected data points`)
    } catch (error) {
      console.error('Error updating data selection:', error)
      toast.error('Error updating data selection. Please try again.')
      setHasError(true)
    }
  }, [calculateStats, analyzeData, setUploadedData, setInsights])

  // Toggle row selection for a specific column
  const toggleRowSelection = useCallback((columnIndex: number, rowIndex: number) => {
    try {
      if (!safeUploadedData || safeUploadedData.length === 0 || columnIndex >= safeUploadedData.length || columnIndex < 0) {
        console.warn('Invalid column index or no data available')
        return
      }
      
      const updatedData = [...safeUploadedData]
      const column = updatedData[columnIndex]
      
      if (column && column.selectedRows && rowIndex >= 0 && rowIndex < column.selectedRows.length) {
        column.selectedRows[rowIndex] = !column.selectedRows[rowIndex]
        updateDataSelection(updatedData)
      }
    } catch (error) {
      console.error('Error toggling row selection:', error)
      setHasError(true)
    }
  }, [safeUploadedData, updateDataSelection])

  // Toggle all rows for a column
  const toggleAllRows = useCallback((columnIndex: number, selectAll: boolean) => {
    if (!safeUploadedData || safeUploadedData.length === 0 || columnIndex >= safeUploadedData.length || columnIndex < 0) {
      console.warn('Invalid column index or no data available')
      return
    }
    
    const updatedData = [...safeUploadedData]
    const column = updatedData[columnIndex]
    
    if (column && column.values) {
      column.selectedRows = new Array(column.values.length).fill(selectAll)
      updateDataSelection(updatedData)
    }
  }, [safeUploadedData, updateDataSelection])

  // Bulk selection by range
  const selectRowRange = useCallback((columnIndex: number, startRow: number, endRow: number, selected: boolean) => {
    if (!safeUploadedData || safeUploadedData.length === 0 || columnIndex >= safeUploadedData.length || columnIndex < 0) {
      return
    }
    
    const updatedData = [...safeUploadedData]
    const column = updatedData[columnIndex]
    
    if (column && column.selectedRows) {
      for (let i = startRow; i <= endRow && i < column.selectedRows.length; i++) {
        column.selectedRows[i] = selected
      }
      updateDataSelection(updatedData)
    }
  }, [safeUploadedData, updateDataSelection])

  // Smart selection presets
  const applySmartSelection = useCallback((type: 'recent' | 'outliers' | 'complete') => {
    if (!safeUploadedData || safeUploadedData.length === 0) {
      console.warn('No data available for smart selection')
      return
    }
    
    // Clear existing filters first
    setFilterConfig({})
    
    const updatedData = [...safeUploadedData]
    
    updatedData.forEach((column, colIndex) => {
      if (!column || !column.values) {
        return // Skip invalid columns
      }
      
      if (type === 'recent' && column.type === 'date' && column.dateValues) {
        // Select most recent 50% of data
        const validDateIndices = column.dateValues
          .map((date, index) => date ? { date, index } : null)
          .filter((item): item is NonNullable<typeof item> => item !== null)
        
        const sortedIndices = validDateIndices
          .sort((a, b) => b.date.getTime() - a.date.getTime())
          .slice(0, Math.ceil(validDateIndices.length * 0.5))
          .map(item => item.index)
        
        column.selectedRows = new Array(column.values.length).fill(false)
        sortedIndices.forEach(index => {
          column.selectedRows![index] = true
        })
      } else if (type === 'outliers' && column.type === 'number' && column.stats) {
        // Use advanced outlier detection
        const numericValues = column.values
          .map(val => Number(val))
          .filter(val => !isNaN(val) && isFinite(val))
        
        if (numericValues.length >= 4) {
          const outlierFlags = detectOutliers(numericValues, 'iqr', 2)
          
          let outlierIndex = 0
          column.selectedRows = column.values.map((val) => {
            const numVal = Number(val)
            if (isNaN(numVal) || !isFinite(numVal)) return false
            
            return outlierFlags[outlierIndex++]
          })
        } else {
          column.selectedRows = new Array(column.values.length).fill(false)
        }
      } else if (type === 'complete') {
        // Select only rows with complete data (no null/empty values)
        column.selectedRows = column.values.map(val => 
          val !== null && val !== undefined && val !== ''
        )
      } else {
        // For non-matching types, keep current selection or select all
        column.selectedRows = column.selectedRows || new Array(column.values.length).fill(true)
      }
    })
    
    updateDataSelection(updatedData)
    
    const selectionType = type === 'recent' ? 'recent data' : 
                         type === 'outliers' ? 'outlier values' : 'complete records'
    toast.success(`Applied ${selectionType} selection`)
  }, [safeUploadedData, detectOutliers, updateDataSelection, setFilterConfig])

  const generateChart = (column: DataColumn, chartType: 'bar' | 'pie' | 'line' = 'bar') => {
    // Add safety checks for column parameter
    if (!column || !column.selectedRows || !column.values) {
      return (
        <div className="h-48 flex items-center justify-center text-muted-foreground">
          Invalid column data
        </div>
      )
    }
    
    // Filter data based on selected rows
    const selectedIndices = column.selectedRows
      ?.map((selected, index) => selected ? index : -1)
      .filter(index => index !== -1) || []
    
    if (selectedIndices.length === 0) {
      return (
        <div className="h-48 flex items-center justify-center text-muted-foreground">
          No data points selected for visualization
        </div>
      )
    }

    if (column.type === 'date') {
      // Create time series chart for selected date columns
      if (!column.dateValues) return null
      
      const selectedTimeSeriesData = selectedIndices
        .map(index => {
          const date = column.dateValues![index]
          return (date && date instanceof Date && !isNaN(date.getTime())) ? {
            date: date.toLocaleDateString(),
            value: 1,
            originalDate: date
          } : null
        })
        .filter((item): item is NonNullable<typeof item> => item !== null)
        .sort((a, b) => a.originalDate.getTime() - b.originalDate.getTime())
      
      // Group by time period if too many points
      let groupedData = selectedTimeSeriesData
      if (selectedTimeSeriesData.length > 20) {
        const groups = new Map<string, number>()
        selectedTimeSeriesData.forEach(item => {
          const key = item.originalDate.toISOString().substring(0, 7) // Group by month
          groups.set(key, (groups.get(key) || 0) + 1)
        })
        
        groupedData = Array.from(groups.entries()).map(([key, count]) => ({
          date: new Date(key).toLocaleDateString('en-US', { year: 'numeric', month: 'short' }),
          value: count,
          originalDate: new Date(key)
        }))
      }
      
      return (
        <ResponsiveContainer width="100%" height={200}>
          <LineChart data={groupedData}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="date" />
            <YAxis />
            <Line type="monotone" dataKey="value" stroke={CHART_COLORS[2]} strokeWidth={2} />
          </LineChart>
        </ResponsiveContainer>
      )
    }
    
    if (column.type !== 'number') return null

    // Prepare data for charts using selected indices
    const chartData = selectedIndices.map((index, chartIndex) => ({
      name: `Item ${chartIndex + 1}`,
      value: Number(column.values[index])
    })).slice(0, 10) // Limit to first 10 items for readability

    if (chartType === 'bar') {
      return (
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
            <defs>
              <linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="var(--color-primary)" stopOpacity={0.8}/>
                <stop offset="95%" stopColor="var(--color-primary)" stopOpacity={0.1}/>
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="var(--color-border)" />
            <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fill: "var(--color-muted-foreground)", fontSize: 12}} dy={10} />
            <YAxis axisLine={false} tickLine={false} tick={{fill: "var(--color-muted-foreground)", fontSize: 12}} />
            <Tooltip content={<CustomTooltip />} />
            <Bar dataKey="value" fill="url(#colorValue)" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      )
    } else if (chartType === 'pie') {
      return (
        <ResponsiveContainer width="100%" height={200}>
          <RechartsPieChart>
            <Tooltip content={<CustomTooltip />} />
            <Pie 
              data={chartData} 
              dataKey="value" 
              nameKey="name" 
              cx="50%" 
              cy="50%" 
              outerRadius={80} 
              innerRadius={60}
              paddingAngle={5}
              stroke="none"
            >
              {chartData.map((_, index) => (
                <Cell key={`cell-${index}`} fill={`var(--color-chart-${(index % 5) + 1})`} />
              ))}
            </Pie>
          </RechartsPieChart>
        </ResponsiveContainer>
      )
    }

    return null
  }
  
  // Generate time series chart combining date and numeric columns with selected data
  const generateTimeSeriesChart = (dateColumn: DataColumn, numericColumn: DataColumn) => {
    // Add safety checks for both columns
    if (!dateColumn || !numericColumn || 
        !dateColumn.dateValues || !dateColumn.selectedRows || !dateColumn.values ||
        !numericColumn.selectedRows || !numericColumn.values ||
        dateColumn.values.length !== numericColumn.values.length) {
      return (
        <div className="h-60 flex items-center justify-center text-muted-foreground">
          Invalid or mismatched column data
        </div>
      )
    }
    
    // Get indices of rows selected in both columns
    const selectedIndices = dateColumn.selectedRows
      ?.map((selected, index) => (selected && numericColumn.selectedRows?.[index]) ? index : -1)
      .filter(index => index !== -1) || []
    
    if (selectedIndices.length === 0) {
      return (
        <div className="h-60 flex items-center justify-center text-muted-foreground">
          No overlapping selected data points for time series visualization
        </div>
      )
    }
    
    const timeSeriesData: TimeSeriesData[] = []
    selectedIndices.forEach(index => {
      const date = dateColumn.dateValues![index]
      const value = Number(numericColumn.values[index])
      if (date && date instanceof Date && !isNaN(date.getTime()) && !isNaN(value)) {
        timeSeriesData.push({
          date: date.toLocaleDateString(),
          value,
          originalDate: date
        })
      }
    })
    
    // Sort by date
    timeSeriesData.sort((a, b) => a.originalDate.getTime() - b.originalDate.getTime())
    
    return (
      <ResponsiveContainer width="100%" height={250}>
        <LineChart data={timeSeriesData}>
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey="date" />
          <YAxis />
          <Line 
            type="monotone" 
            dataKey="value" 
            stroke={CHART_COLORS[1]} 
            strokeWidth={2}
            dot={{ fill: CHART_COLORS[1], strokeWidth: 2, r: 4 }}
          />
        </LineChart>
      </ResponsiveContainer>
    )
  }

  const exportInfographic = useCallback(async () => {
    const element = document.getElementById('dashboard-container')
    if (!element) return
    
    try {
      toast.info('Generating PDF report...', { id: 'pdf-export' })
      // Use html2canvas to capture the visible dashboard area
      const canvas = await html2canvas(element, { scale: 2, useCORS: true, backgroundColor: '#ffffff' })
      const imgData = canvas.toDataURL('image/png')
      
      const pdf = new jsPDF('p', 'mm', 'a4')
      const pdfWidth = pdf.internal.pageSize.getWidth()
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width
      
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight)
      pdf.save(`${fileName.split('.')[0] || 'analysis'}-report.pdf`)
      
      toast.success('Report exported successfully!', { id: 'pdf-export' })
    } catch (error) {
      console.error('Error generating PDF:', error)
      toast.error('Failed to generate PDF', { id: 'pdf-export' })
    }
  }, [fileName])

  const clearData = useCallback(() => {
    setUploadedData([])
    setInsights([])
    setFileName('')
    setFilterConfig({})
    toast.success('Data cleared successfully!')
  }, [setUploadedData, setInsights, setFileName, setFilterConfig])

  if (safeUploadedData.length === 0) {
    return (
      <div className="min-h-screen bg-background p-6">
        <div className="max-w-4xl mx-auto">
          <div className="text-center mb-12">
            <h1 className="text-4xl font-bold text-foreground mb-4">
              Excel Data Analysis & Infographic Generator
            </h1>
            <p className="text-xl text-muted-foreground">
              Transform your Excel data into beautiful, shareable infographics with automated insights
            </p>
          </div>

          <Card className="max-w-2xl mx-auto">
            <CardHeader className="text-center">
              <CardTitle className="flex items-center justify-center gap-2">
                <FileSpreadsheet className="w-8 h-8 text-primary" />
                Upload Your Excel File
              </CardTitle>
              <CardDescription>
                Upload an Excel file (.xlsx, .xls) or CSV file to get started. Maximum file size: 10MB
              </CardDescription>
            </CardHeader>
            <CardContent>
              {isProcessing ? (
                <div className="space-y-4">
                  <Progress value={33} className="w-full" />
                  <p className="text-center text-muted-foreground">{processingStep}</p>
                </div>
              ) : (
                <div className="space-y-6">
                  <div className="border-2 border-dashed border-border rounded-lg p-12 text-center hover:border-primary/50 transition-colors">
                    <Upload className="w-12 h-12 text-muted-foreground mx-auto mb-4" />
                    <label className="cursor-pointer">
                      <span className="text-lg font-medium text-foreground">
                        Click to upload or drag and drop
                      </span>
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        onChange={handleFileUpload}
                        className="hidden"
                      />
                    </label>
                    <p className="text-muted-foreground mt-2">
                      Supports Excel (.xlsx, .xls) and CSV files
                    </p>
                    <p className="text-xs text-muted-foreground mt-2">
                      Date formats: YYYY-MM-DD, MM/DD/YYYY, DD.MM.YYYY, Month DD YYYY, Excel serial dates, and more
                    </p>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="text-center p-4">
                      <Upload className="w-8 h-8 text-primary mx-auto mb-2" />
                      <h3 className="font-medium">1. Upload</h3>
                      <p className="text-sm text-muted-foreground">Upload Excel or CSV with any date format</p>
                    </div>
                    <div className="text-center p-4">
                      <TrendingUp className="w-8 h-8 text-primary mx-auto mb-2" />
                      <h3 className="font-medium">2. Auto-Detect</h3>
                      <p className="text-sm text-muted-foreground">Smart parsing of mixed date formats</p>
                    </div>
                    <div className="text-center p-4">
                      <Download className="w-8 h-8 text-primary mx-auto mb-2" />
                      <h3 className="font-medium">3. Analyze</h3>
                      <p className="text-sm text-muted-foreground">Get insights with time-series analysis</p>
                    </div>
                  </div>
                </div>
              )}
            </CardContent>
          </Card>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-background p-6">
      <div className="max-w-6xl mx-auto" id="dashboard-container">
        <div className="flex items-center justify-between mb-8">
          <div>
            <h1 className="text-3xl font-bold text-foreground">Data Analysis Results</h1>
            <p className="text-muted-foreground">File: {fileName}</p>
          </div>
          <div className="flex gap-2">
            <Button onClick={exportInfographic} className="flex items-center gap-2">
              <Download className="w-4 h-4" />
              Export Report
            </Button>
            <Button variant="outline" onClick={clearData}>
              Upload New File
            </Button>
          </div>
        </div>

        <Tabs defaultValue="insights" className="space-y-6">
          <TabsList className="grid w-full grid-cols-6">
            <TabsTrigger value="insights" className="flex items-center gap-2">
              <TrendingUp className="w-4 h-4" />
              Insights
            </TabsTrigger>
            <TabsTrigger value="selection" className="flex items-center gap-2">
              <Funnel className="w-4 h-4" />
              Data Selection
            </TabsTrigger>
            <TabsTrigger value="timeseries" className="flex items-center gap-2">
              <Clock className="w-4 h-4" />
              Time Series
            </TabsTrigger>
            <TabsTrigger value="charts" className="flex items-center gap-2">
              <BarChart3 className="w-4 h-4" />
              Charts
            </TabsTrigger>
            <TabsTrigger value="data" className="flex items-center gap-2">
              <FileSpreadsheet className="w-4 h-4" />
              Data
            </TabsTrigger>
          
            <TabsTrigger value="builder" className="flex items-center gap-2">
              <Wrench className="w-4 h-4" />
              Builder
            </TabsTrigger>
          </TabsList>

          <TabsContent value="insights" className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {safeInsights.map((insight, index) => (
                <Card key={index}>
                  <CardHeader>
                    <div className="flex items-start justify-between">
                      <CardTitle className="text-lg">{insight.title}</CardTitle>
                      <Badge variant={insight.importance === 'high' ? 'default' : 'secondary'}>
                        {insight.importance}
                      </Badge>
                    </div>
                  </CardHeader>
                  <CardContent>
                    <p className="text-muted-foreground mb-3">{insight.description}</p>
                    {insight.value && (
                      <div className="text-2xl font-bold text-primary">{insight.value}</div>
                    )}
                  </CardContent>
                </Card>
              ))}
            </div>
          </TabsContent>

          <TabsContent value="selection" className="space-y-6">
            <div className="flex flex-col gap-6">
              {/* Selection Overview and Controls */}
              <Card>
                <CardHeader>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <Funnel className="w-5 h-5" />
                      <CardTitle>Data Point Selection</CardTitle>
                    </div>
                    <div className="flex items-center gap-2">
                      <Badge variant="secondary">
                        {(() => {
                          const totalSelected = safeUploadedData.reduce((sum, col) => 
                            sum + (col?.selectedRows?.filter(Boolean).length || 0), 0
                          )
                          const totalPoints = safeUploadedData.reduce((sum, col) => 
                            sum + (col?.values?.length || 0), 0
                          )
                          return `${totalSelected}/${totalPoints} points selected`
                        })()}
                      </Badge>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => {
                          const updatedData = safeUploadedData.map(col => ({
                            ...col,
                            selectedRows: new Array(col.values?.length || 0).fill(true)
                          }))
                          updateDataSelection(updatedData)
                        }}
                      >
                        Select All
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => {
                          const updatedData = safeUploadedData.map(col => ({
                            ...col,
                            selectedRows: new Array(col.values?.length || 0).fill(false)
                          }))
                          updateDataSelection(updatedData)
                        }}
                      >
                        Clear All
                      </Button>
                    </div>
                  </div>
                  <CardDescription>
                    Select specific data points to focus your analysis. Only selected points will be used for statistics and insights generation.
                  </CardDescription>
                </CardHeader>
              </Card>

              {/* Advanced Filters Section */}
              <Card>
                <Collapsible open={showAdvancedFilters} onOpenChange={setShowAdvancedFilters}>
                  <CollapsibleTrigger asChild>
                    <CardHeader className="hover:bg-muted/50 cursor-pointer">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <Filter className="w-5 h-5" />
                          <CardTitle>Advanced Filters</CardTitle>
                        </div>
                        <div className="flex items-center gap-2">
                          {Object.keys(safeFilterConfig).length > 0 && (
                            <Badge variant="secondary">
                              {Object.keys(safeFilterConfig).length} active
                            </Badge>
                          )}
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={(e) => {
                              e.stopPropagation()
                              clearAllFilters()
                            }}
                          >
                            Clear All
                          </Button>
                        </div>
                      </div>
                      <CardDescription>
                        Apply sophisticated filters including date ranges, outlier detection, and value constraints
                      </CardDescription>
                    </CardHeader>
                  </CollapsibleTrigger>
                  <CollapsibleContent>
                    <CardContent className="space-y-6">
                      {/* Date Range Filter */}
                      {safeUploadedData.some(col => col?.type === 'date') && (
                        <div className="space-y-4">
                          <div className="flex items-center gap-2">
                            <CalendarBlank className="w-4 h-4" />
                            <Label className="text-sm font-medium">Date Range Filter</Label>
                          </div>
                          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                            <Select
                              value={safeFilterConfig.dateRange?.columnIndex?.toString() || ''}
                              onValueChange={(value) => {
                                const columnIndex = parseInt(value)
                                const column = safeUploadedData[columnIndex]
                                if (column?.type === 'date' && column.stats) {
                                  const minDate = column.stats.min instanceof Date ? column.stats.min : new Date()
                                  const maxDate = column.stats.max instanceof Date ? column.stats.max : new Date()
                                  updateFilterConfig({
                                    dateRange: {
                                      columnIndex,
                                      start: minDate.toISOString().split('T')[0],
                                      end: maxDate.toISOString().split('T')[0]
                                    }
                                  })
                                }
                              }}
                            >
                              <SelectTrigger>
                                <SelectValue placeholder="Select date column" />
                              </SelectTrigger>
                              <SelectContent>
                                {safeUploadedData
                                  .map((col, index) => col?.type === 'date' ? { col, index } : null)
                                  .filter(Boolean)
                                  .map(({ col, index }) => (
                                    <SelectItem key={index} value={index.toString()}>
                                      {col!.name}
                                    </SelectItem>
                                  ))
                                }
                              </SelectContent>
                            </Select>
                            <div className="space-y-1">
                              <Label className="text-xs">Start Date</Label>
                              <Input
                                type="date"
                                value={safeFilterConfig.dateRange?.start || ''}
                                onChange={(e) => {
                                  if (safeFilterConfig.dateRange) {
                                    updateFilterConfig({
                                      dateRange: {
                                        ...safeFilterConfig.dateRange,
                                        start: e.target.value
                                      }
                                    })
                                  }
                                }}
                                disabled={!safeFilterConfig.dateRange}
                              />
                            </div>
                            <div className="space-y-1">
                              <Label className="text-xs">End Date</Label>
                              <Input
                                type="date"
                                value={safeFilterConfig.dateRange?.end || ''}
                                onChange={(e) => {
                                  if (safeFilterConfig.dateRange) {
                                    updateFilterConfig({
                                      dateRange: {
                                        ...safeFilterConfig.dateRange,
                                        end: e.target.value
                                      }
                                    })
                                  }
                                }}
                                disabled={!safeFilterConfig.dateRange}
                              />
                            </div>
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() => {
                                const { dateRange, ...rest } = safeFilterConfig
                                updateFilterConfig(rest)
                              }}
                              disabled={!safeFilterConfig.dateRange}
                            >
                              <X className="w-3 h-3 mr-1" />
                              Clear
                            </Button>
                          </div>
                        </div>
                      )}

                      {/* Outlier Detection Filter */}
                      {safeUploadedData.some(col => col?.type === 'number') && (
                        <div className="space-y-4">
                          <div className="flex items-center gap-2">
                            <TrendUp className="w-4 h-4" />
                            <Label className="text-sm font-medium">Statistical Outlier Detection</Label>
                          </div>
                          <div className="grid grid-cols-1 md:grid-cols-6 gap-4">
                            <Select
                              value={safeFilterConfig.outlierDetection?.columnIndex?.toString() || ''}
                              onValueChange={(value) => {
                                const columnIndex = parseInt(value)
                                updateFilterConfig({
                                  outlierDetection: {
                                    columnIndex,
                                    sensitivity: 2,
                                    method: 'iqr',
                                    selectOutliers: true
                                  }
                                })
                              }}
                            >
                              <SelectTrigger>
                                <SelectValue placeholder="Select numeric column" />
                              </SelectTrigger>
                              <SelectContent>
                                {safeUploadedData
                                  .map((col, index) => col?.type === 'number' ? { col, index } : null)
                                  .filter(Boolean)
                                  .map(({ col, index }) => (
                                    <SelectItem key={index} value={index.toString()}>
                                      {col!.name}
                                    </SelectItem>
                                  ))
                                }
                              </SelectContent>
                            </Select>
                            <Select
                              value={safeFilterConfig.outlierDetection?.method || 'iqr'}
                              onValueChange={(method: 'iqr' | 'zscore' | 'modified_zscore') => {
                                if (safeFilterConfig.outlierDetection) {
                                  updateFilterConfig({
                                    outlierDetection: {
                                      ...safeFilterConfig.outlierDetection,
                                      method
                                    }
                                  })
                                }
                              }}
                              disabled={!safeFilterConfig.outlierDetection}
                            >
                              <SelectTrigger>
                                <SelectValue placeholder="Method" />
                              </SelectTrigger>
                              <SelectContent>
                                <SelectItem value="iqr">IQR Method</SelectItem>
                                <SelectItem value="zscore">Z-Score</SelectItem>
                                <SelectItem value="modified_zscore">Modified Z-Score</SelectItem>
                              </SelectContent>
                            </Select>
                            <div className="space-y-1">
                              <Label className="text-xs">Sensitivity</Label>
                              <div className="px-3">
                                <Slider
                                  value={[safeFilterConfig.outlierDetection?.sensitivity || 2]}
                                  onValueChange={([sensitivity]) => {
                                    if (safeFilterConfig.outlierDetection) {
                                      updateFilterConfig({
                                        outlierDetection: {
                                          ...safeFilterConfig.outlierDetection,
                                          sensitivity
                                        }
                                      })
                                    }
                                  }}
                                  min={1}
                                  max={3}
                                  step={1}
                                  disabled={!safeFilterConfig.outlierDetection}
                                />
                                <div className="flex justify-between text-xs text-muted-foreground mt-1">
                                  <span>Conservative</span>
                                  <span>Aggressive</span>
                                </div>
                              </div>
                            </div>
                            <div className="flex items-center space-x-2 pt-6">
                              <Checkbox
                                id="select-outliers"
                                checked={safeFilterConfig.outlierDetection?.selectOutliers ?? true}
                                onCheckedChange={(checked) => {
                                  if (safeFilterConfig.outlierDetection) {
                                    updateFilterConfig({
                                      outlierDetection: {
                                        ...safeFilterConfig.outlierDetection,
                                        selectOutliers: checked as boolean
                                      }
                                    })
                                  }
                                }}
                                disabled={!safeFilterConfig.outlierDetection}
                              />
                              <Label htmlFor="select-outliers" className="text-xs">
                                Select outliers
                              </Label>
                            </div>
                            <div className="text-xs text-muted-foreground">
                              {safeFilterConfig.outlierDetection && (() => {
                                const count = safeUploadedData[safeFilterConfig.outlierDetection.columnIndex]?.selectedRows?.filter(Boolean).length || 0
                                return `${count} data points selected`
                              })()}
                            </div>
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() => {
                                const { outlierDetection, ...rest } = safeFilterConfig
                                updateFilterConfig(rest)
                              }}
                              disabled={!safeFilterConfig.outlierDetection}
                            >
                              <X className="w-3 h-3 mr-1" />
                              Clear
                            </Button>
                          </div>
                        </div>
                      )}

                      {/* Value Range Filter */}
                      {safeUploadedData.some(col => col?.type === 'number') && (
                        <div className="space-y-4">
                          <div className="flex items-center gap-2">
                            <TrendDown className="w-4 h-4" />
                            <Label className="text-sm font-medium">Value Range Filter</Label>
                          </div>
                          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                            <Select
                              value={safeFilterConfig.valueRange?.columnIndex?.toString() || ''}
                              onValueChange={(value) => {
                                const columnIndex = parseInt(value)
                                const column = safeUploadedData[columnIndex]
                                if (column?.type === 'number' && column.stats) {
                                  updateFilterConfig({
                                    valueRange: {
                                      columnIndex,
                                      min: column.stats.min as number || 0,
                                      max: column.stats.max as number || 100
                                    }
                                  })
                                }
                              }}
                            >
                              <SelectTrigger>
                                <SelectValue placeholder="Select numeric column" />
                              </SelectTrigger>
                              <SelectContent>
                                {safeUploadedData
                                  .map((col, index) => col?.type === 'number' ? { col, index } : null)
                                  .filter(Boolean)
                                  .map(({ col, index }) => (
                                    <SelectItem key={index} value={index.toString()}>
                                      {col!.name}
                                    </SelectItem>
                                  ))
                                }
                              </SelectContent>
                            </Select>
                            <div className="space-y-1">
                              <Label className="text-xs">Min Value</Label>
                              <Input
                                type="number"
                                value={safeFilterConfig.valueRange?.min || ''}
                                onChange={(e) => {
                                  if (safeFilterConfig.valueRange) {
                                    updateFilterConfig({
                                      valueRange: {
                                        ...safeFilterConfig.valueRange,
                                        min: parseFloat(e.target.value) || 0
                                      }
                                    })
                                  }
                                }}
                                disabled={!safeFilterConfig.valueRange}
                              />
                            </div>
                            <div className="space-y-1">
                              <Label className="text-xs">Max Value</Label>
                              <Input
                                type="number"
                                value={safeFilterConfig.valueRange?.max || ''}
                                onChange={(e) => {
                                  if (safeFilterConfig.valueRange) {
                                    updateFilterConfig({
                                      valueRange: {
                                        ...safeFilterConfig.valueRange,
                                        max: parseFloat(e.target.value) || 100
                                      }
                                    })
                                  }
                                }}
                                disabled={!safeFilterConfig.valueRange}
                              />
                            </div>
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() => {
                                const { valueRange, ...rest } = safeFilterConfig
                                updateFilterConfig(rest)
                              }}
                              disabled={!safeFilterConfig.valueRange}
                            >
                              <X className="w-3 h-3 mr-1" />
                              Clear
                            </Button>
                          </div>
                        </div>
                      )}

                      {/* Text Pattern Filter */}
                      <div className="space-y-4">
                        <div className="flex items-center gap-2">
                          <FileSpreadsheet className="w-4 h-4" />
                          <Label className="text-sm font-medium">Text Pattern Filter</Label>
                        </div>
                        <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
                          <Select
                            value={safeFilterConfig.textFilter?.columnIndex?.toString() || ''}
                            onValueChange={(value) => {
                              const columnIndex = parseInt(value)
                              updateFilterConfig({
                                textFilter: {
                                  columnIndex,
                                  pattern: '',
                                  caseSensitive: false
                                }
                              })
                            }}
                          >
                            <SelectTrigger>
                              <SelectValue placeholder="Select column" />
                            </SelectTrigger>
                            <SelectContent>
                              {safeUploadedData.map((col, index) => (
                                <SelectItem key={index} value={index.toString()}>
                                  {col?.name || `Column ${index + 1}`}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                          <div className="space-y-1">
                            <Label className="text-xs">Search Pattern</Label>
                            <Input
                              type="text"
                              placeholder="Enter text to search"
                              value={safeFilterConfig.textFilter?.pattern || ''}
                              onChange={(e) => {
                                if (safeFilterConfig.textFilter) {
                                  updateFilterConfig({
                                    textFilter: {
                                      ...safeFilterConfig.textFilter,
                                      pattern: e.target.value
                                    }
                                  })
                                }
                              }}
                              disabled={!safeFilterConfig.textFilter}
                            />
                          </div>
                          <div className="flex items-center space-x-2 pt-6">
                            <Checkbox
                              id="case-sensitive"
                              checked={safeFilterConfig.textFilter?.caseSensitive || false}
                              onCheckedChange={(checked) => {
                                if (safeFilterConfig.textFilter) {
                                  updateFilterConfig({
                                    textFilter: {
                                      ...safeFilterConfig.textFilter,
                                      caseSensitive: checked as boolean
                                    }
                                  })
                                }
                              }}
                              disabled={!safeFilterConfig.textFilter}
                            />
                            <Label htmlFor="case-sensitive" className="text-xs">
                              Case sensitive
                            </Label>
                          </div>
                          <div className="text-xs text-muted-foreground pt-6">
                            {safeFilterConfig.textFilter && (() => {
                              const count = safeUploadedData[safeFilterConfig.textFilter.columnIndex]?.selectedRows?.filter(Boolean).length || 0
                              return `${count} matches`
                            })()}
                          </div>
                          <Button
                            variant="outline"
                            size="sm"
                            onClick={() => {
                              const { textFilter, ...rest } = safeFilterConfig
                              updateFilterConfig(rest)
                            }}
                            disabled={!safeFilterConfig.textFilter}
                          >
                            <X className="w-3 h-3 mr-1" />
                            Clear
                          </Button>
                        </div>
                      </div>
                    </CardContent>
                  </CollapsibleContent>
                </Collapsible>
              </Card>
              {/* Smart Selection Presets */}
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <Funnel className="w-5 h-5" />
                    Smart Selection Presets
                  </CardTitle>
                  <CardDescription>
                    Apply intelligent filters to focus your analysis on specific data patterns
                  </CardDescription>
                </CardHeader>
                <CardContent>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                    <Button 
                      variant="outline" 
                      onClick={() => applySmartSelection('recent')}
                      className="flex items-center gap-2"
                    >
                      <Calendar className="w-4 h-4" />
                      Recent Data (50%)
                    </Button>
                    <Button 
                      variant="outline" 
                      onClick={() => applySmartSelection('outliers')}
                      className="flex items-center gap-2"
                    >
                      <TrendingUp className="w-4 h-4" />
                      Outlier Values
                    </Button>
                    <Button 
                      variant="outline" 
                      onClick={() => applySmartSelection('complete')}
                      className="flex items-center gap-2"
                    >
                      <Check className="w-4 h-4" />
                      Complete Records
                    </Button>
                  </div>
                </CardContent>
              </Card>

              {/* Column Selection Overview */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {safeUploadedData.filter(column => column && column.values && column.selectedRows).map((column, columnIndex) => {
                  const selectedCount = column.selectedRows?.filter(Boolean).length || 0
                  const totalCount = column.values.length
                  const selectionPercentage = Math.round((selectedCount / totalCount) * 100)
                  
                  return (
                    <Card key={columnIndex}>
                      <CardHeader>
                        <div className="flex items-center justify-between">
                          <CardTitle className="text-lg">{column.name}</CardTitle>
                          <Badge variant={column.type === 'number' ? 'default' : column.type === 'date' ? 'secondary' : 'outline'}>
                            {column.type}
                          </Badge>
                        </div>
                        <CardDescription>
                          {selectedCount}/{totalCount} selected ({selectionPercentage}%)
                        </CardDescription>
                      </CardHeader>
                      <CardContent className="space-y-4">
                        {/* Column-level controls */}
                        <div className="grid grid-cols-2 gap-2 mb-4">
                          <div className="grid grid-cols-2 gap-1">
                            <Button 
                              size="sm" 
                              variant="outline"
                              onClick={() => toggleAllRows(columnIndex, true)}
                              className="text-xs"
                            >
                              <Check className="w-3 h-3 mr-1" />
                              All
                            </Button>
                            <Button 
                              size="sm" 
                              variant="outline"
                              onClick={() => toggleAllRows(columnIndex, false)}
                              className="text-xs"
                            >
                              <X className="w-3 h-3 mr-1" />
                              None
                            </Button>
                          </div>
                          <div className="grid grid-cols-2 gap-1">
                            <Button 
                              size="sm" 
                              variant="outline"
                              onClick={() => selectRowRange(columnIndex, 0, Math.floor(column.values.length / 2) - 1, true)}
                              className="text-xs"
                            >
                              First 50%
                            </Button>
                            <Button 
                              size="sm" 
                              variant="outline"
                              onClick={() => selectRowRange(columnIndex, Math.floor(column.values.length / 2), column.values.length - 1, true)}
                              className="text-xs"
                            >
                              Last 50%
                            </Button>
                          </div>
                        </div>
                        
                        <Progress value={selectionPercentage} className="w-full" />
                        
                        {/* Sample data preview with checkboxes */}
                        <div className="space-y-2">
                          <div className="flex items-center justify-between">
                            <Label className="text-sm font-medium">Sample Data (First 10 rows):</Label>
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => setShowDataSelection(!showDataSelection)}
                              className="text-xs"
                            >
                              {showDataSelection ? 'Hide All' : 'Show All'} Rows
                            </Button>
                          </div>
                          <div className="space-y-1 max-h-64 overflow-y-auto">
                            {(showDataSelection ? column.values : column.values.slice(0, 10)).map((value, rowIndex) => (
                              <div key={rowIndex} className="flex items-center space-x-2">
                                <Checkbox
                                  id={`${columnIndex}-${rowIndex}`}
                                  checked={column.selectedRows?.[rowIndex] || false}
                                  onCheckedChange={() => toggleRowSelection(columnIndex, rowIndex)}
                                />
                                <Label 
                                  htmlFor={`${columnIndex}-${rowIndex}`}
                                  className="text-xs flex-1 cursor-pointer truncate"
                                >
                                  <span className="text-muted-foreground mr-2">#{rowIndex + 1}:</span>
                                  {column.type === 'date' && column.dateValues?.[rowIndex] && 
                                   column.dateValues[rowIndex] instanceof Date && 
                                   !isNaN(column.dateValues[rowIndex]!.getTime())
                                    ? column.dateValues[rowIndex]!.toLocaleDateString()
                                    : String(value)
                                  }
                                </Label>
                                {column.selectedRows?.[rowIndex] && (
                                  <Check className="w-3 h-3 text-primary" />
                                )}
                              </div>
                            ))}
                            {!showDataSelection && column.values.length > 10 && (
                              <div className="text-xs text-muted-foreground pt-2">
                                ...and {column.values.length - 10} more rows (click "Show All Rows" to see them)
                              </div>
                            )}
                          </div>
                        </div>
                        
                        {/* Quick stats for selected data */}
                        {column.stats && selectedCount > 0 && (
                          <div className="text-xs text-muted-foreground space-y-1">
                            {column.type === 'number' && (
                              <>
                                <div>Selected Avg: {column.stats.avg?.toLocaleString(undefined, { maximumFractionDigits: 2 })}</div>
                                <div>Selected Range: {(column.stats.min as number)?.toLocaleString()} - {(column.stats.max as number)?.toLocaleString()}</div>
                              </>
                            )}
                            {column.type === 'date' && (
                              <>
                                <div>Selected Range: {column.stats.dateRange}</div>
                                <div>Pattern: {column.stats.frequency}</div>
                              </>
                            )}
                          </div>
                        )}
                      </CardContent>
                    </Card>
                  )
                })}
              </div>
            </div>
          </TabsContent>

          <TabsContent value="timeseries" className="space-y-6">
            {(() => {
              const dateColumns = safeUploadedData.filter(col => col && col.type === 'date' && col.values && col.selectedRows)
              const numericColumns = safeUploadedData.filter(col => col && col.type === 'number' && col.values && col.selectedRows)
              
              if (dateColumns.length === 0) {
                return (
                  <Card>
                    <CardContent className="text-center py-12">
                      <Calendar className="w-12 h-12 text-muted-foreground mx-auto mb-4" />
                      <h3 className="text-lg font-medium mb-2">No Date Columns Found</h3>
                      <p className="text-muted-foreground">
                        To view time series analysis, your data needs to contain date columns. 
                        Make sure your dates are in a recognized format (YYYY-MM-DD, MM/DD/YYYY, etc.)
                      </p>
                    </CardContent>
                  </Card>
                )
              }
              
              return (
                <div className="space-y-6">
                  {/* Date column timeline charts */}
                  {dateColumns.map((dateCol, index) => (
                    <Card key={`date-${index}`}>
                      <CardHeader>
                        <CardTitle className="flex items-center gap-2">
                          <Calendar className="w-5 h-5" />
                          {dateCol.name} Timeline
                        </CardTitle>
                        <CardDescription>
                          {dateCol.stats?.frequency} frequency • {dateCol.stats?.dateRange}
                          {dateCol.stats?.formatSummary && (
                            <div className="text-xs mt-1 text-muted-foreground">
                              Formats: {dateCol.stats.formatSummary}
                            </div>
                          )}
                        </CardDescription>
                      </CardHeader>
                      <CardContent>
                        {generateChart(dateCol, 'line')}
                        <div className="mt-4 grid grid-cols-2 gap-4 text-sm">
                          <div>
                            <span className="text-muted-foreground">Pattern: </span>
                            <Badge variant="outline">{dateCol.stats?.frequency}</Badge>
                          </div>
                          <div>
                            <span className="text-muted-foreground">Data Points: </span>
                            <span className="font-medium">{dateCol.dateValues?.filter(d => d !== null).length}</span>
                          </div>
                        </div>
                      </CardContent>
                    </Card>
                  ))}
                  
                  {/* Time series combinations */}
                  {dateColumns.map(dateCol => 
                    numericColumns
                      .filter(numCol => dateCol.values.length === numCol.values.length)
                      .map((numCol, numIndex) => (
                        <Card key={`timeseries-${dateCol.name}-${numCol.name}`}>
                          <CardHeader>
                            <CardTitle className="flex items-center gap-2">
                              <Clock className="w-5 h-5" />
                              {numCol.name} over {dateCol.name}
                            </CardTitle>
                            <CardDescription>
                              Time series analysis showing how {numCol.name} changes over time
                            </CardDescription>
                          </CardHeader>
                          <CardContent>
                            {generateTimeSeriesChart(dateCol, numCol)}
                            <div className="mt-4 grid grid-cols-3 gap-4 text-sm">
                              <div>
                                <span className="text-muted-foreground">Avg Value: </span>
                                <span className="font-medium">
                                  {numCol.stats?.avg?.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                                </span>
                              </div>
                              <div>
                                <span className="text-muted-foreground">Range: </span>
                                <span className="font-medium">
                                  {(numCol.stats?.min as number)?.toLocaleString()} - {(numCol.stats?.max as number)?.toLocaleString()}
                                </span>
                              </div>
                              <div>
                                <span className="text-muted-foreground">Time Span: </span>
                                <span className="font-medium">{dateCol.stats?.dateRange?.split(' to ').length === 2 ? 
                                  `${dateCol.stats.dateRange.split(' to ')[1].split('/')[2]} - ${dateCol.stats.dateRange.split(' to ')[0].split('/')[2]}` : 
                                  'Unknown'}</span>
                              </div>
                            </div>
                          </CardContent>
                        </Card>
                      ))
                  )}
                </div>
              )
            })()}
          </TabsContent>
          <TabsContent value="charts" className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {safeUploadedData.filter(col => col && col.type === 'number' && col.values && col.selectedRows).map((column, index) => (
                <Card key={index}>
                  <CardHeader>
                    <CardTitle className="flex items-center gap-2">
                      <BarChart3 className="w-5 h-5" />
                      {column.name} Distribution
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    {generateChart(column, 'bar')}
                    <div className="mt-4 grid grid-cols-2 gap-4 text-sm">
                      <div>
                        <span className="text-muted-foreground">Average: </span>
                        <span className="font-medium">
                          {column.stats?.avg?.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                        </span>
                      </div>
                      <div>
                        <span className="text-muted-foreground">Max: </span>
                        <span className="font-medium">{(column.stats?.max as number)?.toLocaleString()}</span>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              ))}
              
              {safeUploadedData.filter(col => col && col.type === 'date' && col.values && col.selectedRows).map((column, index) => (
                <Card key={`date-chart-${index}`}>
                  <CardHeader>
                    <CardTitle className="flex items-center gap-2">
                      <Calendar className="w-5 h-5" />
                      {column.name} Distribution
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    {generateChart(column, 'line')}
                    <div className="mt-4 grid grid-cols-2 gap-4 text-sm">
                      <div>
                        <span className="text-muted-foreground">Pattern: </span>
                        <Badge variant="outline">{column.stats?.frequency}</Badge>
                      </div>
                      <div>
                        <span className="text-muted-foreground">Range: </span>
                        <span className="font-medium">{column.stats?.dateRange}</span>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              ))}
            </div>
          </TabsContent>

          <TabsContent value="data" className="space-y-6">
            <DataTable columnsData={safeUploadedData} />
          </TabsContent>
        
          <TabsContent value="builder" className="space-y-6">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Wrench className="w-5 h-5 text-primary" />
                  Custom Chart Builder
                </CardTitle>
                <div className="text-sm text-muted-foreground">
                  Select any two columns to plot against each other and choose your preferred chart type.
                </div>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
                  <div className="space-y-2">
                    <Label>X-Axis Column</Label>
                    <Select 
                      value={customChartConfig.xAxis} 
                      onValueChange={(v) => setCustomChartConfig(prev => ({ ...prev, xAxis: v }))}
                    >
                      <SelectTrigger>
                        <SelectValue placeholder="Select X-Axis" />
                      </SelectTrigger>
                      <SelectContent>
                        {safeUploadedData.map(col => (
                          <SelectItem key={`x-${col.name}`} value={col.name}>{col.name}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  
                  <div className="space-y-2">
                    <Label>Y-Axis Column (Numeric)</Label>
                    <Select 
                      value={customChartConfig.yAxis} 
                      onValueChange={(v) => setCustomChartConfig(prev => ({ ...prev, yAxis: v }))}
                    >
                      <SelectTrigger>
                        <SelectValue placeholder="Select Y-Axis" />
                      </SelectTrigger>
                      <SelectContent>
                        {safeUploadedData.map(col => (
                          <SelectItem key={`y-${col.name}`} value={col.name}>{col.name}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  
                  <div className="space-y-2">
                    <Label>Chart Type</Label>
                    <Select 
                      value={customChartConfig.chartType} 
                      onValueChange={(v) => setCustomChartConfig(prev => ({ ...prev, chartType: v as any }))}
                    >
                      <SelectTrigger>
                        <SelectValue placeholder="Select Chart Type" />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="bar">Bar Chart</SelectItem>
                        <SelectItem value="line">Line Chart</SelectItem>
                        <SelectItem value="area">Area Chart</SelectItem>
                        <SelectItem value="scatter">Scatter Plot</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                </div>

                {customChartConfig.xAxis && customChartConfig.yAxis ? (
                  <div className="h-[400px] w-full border rounded-xl p-4 bg-slate-50/50">
                    <ResponsiveContainer width="100%" height="100%">
                      {customChartConfig.chartType === 'bar' ? (
                        <BarChart data={
                          safeUploadedData.find(c => c.name === customChartConfig.xAxis)?.values.slice(0, 500).map((xVal, i) => ({
                            name: String(xVal),
                            value: Number(safeUploadedData.find(c => c.name === customChartConfig.yAxis)?.values[i]) || 0
                          })) || []
                        }>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="name" fontSize={12} tickLine={false} axisLine={false} />
                          <YAxis fontSize={12} tickLine={false} axisLine={false} />
                          <Tooltip contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                          <Bar dataKey="value" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                        </BarChart>
                      ) : customChartConfig.chartType === 'line' ? (
                        <LineChart data={
                          safeUploadedData.find(c => c.name === customChartConfig.xAxis)?.values.slice(0, 500).map((xVal, i) => ({
                            name: String(xVal),
                            value: Number(safeUploadedData.find(c => c.name === customChartConfig.yAxis)?.values[i]) || 0
                          })) || []
                        }>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="name" fontSize={12} tickLine={false} axisLine={false} />
                          <YAxis fontSize={12} tickLine={false} axisLine={false} />
                          <Tooltip contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                          <Line type="monotone" dataKey="value" stroke="#3b82f6" strokeWidth={3} dot={false} />
                        </LineChart>
                      ) : customChartConfig.chartType === 'area' ? (
                        <AreaChart data={
                          safeUploadedData.find(c => c.name === customChartConfig.xAxis)?.values.slice(0, 500).map((xVal, i) => ({
                            name: String(xVal),
                            value: Number(safeUploadedData.find(c => c.name === customChartConfig.yAxis)?.values[i]) || 0
                          })) || []
                        }>
                          <defs>
                            <linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="name" fontSize={12} tickLine={false} axisLine={false} />
                          <YAxis fontSize={12} tickLine={false} axisLine={false} />
                          <Tooltip contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                          <Area type="monotone" dataKey="value" stroke="#3b82f6" fillOpacity={1} fill="url(#colorValue)" />
                        </AreaChart>
                      ) : (
                        <ScatterChart>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis type="category" dataKey="name" name={customChartConfig.xAxis} fontSize={12} tickLine={false} axisLine={false} />
                          <YAxis type="number" dataKey="value" name={customChartConfig.yAxis} fontSize={12} tickLine={false} axisLine={false} />
                          <Tooltip cursor={{ strokeDasharray: '3 3' }} contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                          <Scatter data={
                            safeUploadedData.find(c => c.name === customChartConfig.xAxis)?.values.slice(0, 500).map((xVal, i) => ({
                              name: String(xVal),
                              value: Number(safeUploadedData.find(c => c.name === customChartConfig.yAxis)?.values[i]) || 0
                            })) || []
                          } fill="#3b82f6" />
                        </ScatterChart>
                      )}
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <div className="h-[400px] w-full border border-dashed rounded-xl flex items-center justify-center bg-slate-50 text-muted-foreground">
                    Select X and Y axes to generate your custom chart
                  </div>
                )}
              </CardContent>
            </Card>
          </TabsContent>
  </Tabs>
      </div>
    </div>
  )
}

export default App