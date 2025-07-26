import React, { useState, useCallback } from 'react'
import { useKV } from '@github/spark/hooks'
import * as XLSX from 'xlsx'
import { Upload, BarChart3, Download, FileSpreadsheet, TrendingUp, PieChart, Calendar, Clock } from '@phosphor-icons/react'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Progress } from '@/components/ui/progress'
import { Badge } from '@/components/ui/badge'
import { Separator } from '@/components/ui/separator'
import { toast } from 'sonner'
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, ResponsiveContainer, PieChart as RechartsPieChart, Cell, LineChart, Line, Pie } from 'recharts'

interface DataColumn {
  name: string
  type: 'number' | 'text' | 'date'
  values: any[]
  dateValues?: Date[] // Parsed dates for date columns
  detectedFormats?: string[] // Formats detected in date columns
  stats?: {
    min?: number | Date
    max?: number | Date
    avg?: number
    sum?: number
    count: number
    dateRange?: string // For date columns
    frequency?: 'daily' | 'weekly' | 'monthly' | 'yearly' | 'irregular'
    formatSummary?: string // Summary of detected date formats
  }
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

const CHART_COLORS = ['#1e40af', '#f97316', '#10b981', '#8b5cf6', '#ef4444', '#06b6d4']

function App() {
  const [uploadedData, setUploadedData] = useKV<DataColumn[]>('excel-data', [])
  const [insights, setInsights] = useKV<Insight[]>('data-insights', [])
  const [isProcessing, setIsProcessing] = useState(false)
  const [processingStep, setProcessingStep] = useState('')
  const [fileName, setFileName] = useKV<string>('file-name', '')

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
      if (dateCol.dateValues.length >= 12 && (frequency === 'monthly' || frequency === 'yearly')) {
        const monthCounts = new Array(12).fill(0)
        dateCol.dateValues.forEach(date => {
          monthCounts[date.getMonth()]++
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
    
    // Find different column types
    const numericColumns = columns.filter(col => col.type === 'number' && col.stats)
    const dateColumns = columns.filter(col => col.type === 'date')
    const textColumns = columns.filter(col => col.type === 'text')
    
    // Always add a basic overview
    const totalRows = columns[0]?.values.length || 0
    newInsights.push({
      type: 'summary',
      title: 'Dataset Overview',
      description: `Analyzed ${totalRows} rows across ${columns.length} columns: ${numericColumns.length} numeric, ${dateColumns.length} date, ${textColumns.length} text`,
      value: `${totalRows} records`,
      importance: 'high'
    })
    
    // Generate time-series insights if date columns exist
    if (dateColumns.length > 0) {
      const timeSeriesInsights = generateTimeSeriesInsights(dateColumns, numericColumns)
      newInsights.push(...timeSeriesInsights)
      
      // Add date format detection summary
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
          title: 'Mixed Date Format Detection',
          description: `Successfully detected and unified ${uniqueFormats.length} different date formats: ${formatsList}. All dates normalized for consistent analysis.`,
          value: `${uniqueFormats.length} formats`,
          importance: 'high'
        })
      } else if (formatsList) {
        newInsights.push({
          type: 'summary',
          title: 'Date Format Detection',
          description: `Consistent date format detected: ${formatsList}`,
          importance: 'medium'
        })
      }
    }
    
    if (numericColumns.length > 0) {
      // Find highest value column
      const maxColumn = numericColumns.reduce((max, col) => 
        (col.stats!.max as number || 0) > (max.stats!.max as number || 0) ? col : max
      )
      
      newInsights.push({
        type: 'trend',
        title: 'Highest Values',
        description: `${maxColumn.name} shows the highest peak value in your dataset`,
        value: (maxColumn.stats!.max as number)?.toLocaleString(),
        importance: 'high'
      })

      // Find average insights
      numericColumns.forEach(col => {
        if (col.stats!.avg && col.stats!.avg > 0) {
          newInsights.push({
            type: 'summary',
            title: `${col.name} Average`,
            description: `The average ${col.name.toLowerCase()} across all records`,
            value: col.stats!.avg.toLocaleString(undefined, { maximumFractionDigits: 2 }),
            importance: 'medium'
          })
        }
      })

      // Distribution insight
      if (numericColumns.length >= 2) {
        newInsights.push({
          type: 'correlation',
          title: 'Data Distribution',
          description: `Your data spans multiple metrics, enabling comprehensive analysis across ${numericColumns.length} dimensions`,
          importance: 'medium'
        })
      }
    } else if (textColumns.length > 0) {
      // Add insights for text-based data
      newInsights.push({
        type: 'summary',
        title: 'Text Data Detected',
        description: `Your dataset contains ${textColumns.length} text columns. Consider formatting numeric data as numbers for statistical analysis.`,
        importance: 'medium'
      })
      
      // Show column types
      newInsights.push({
        type: 'summary',
        title: 'Column Types',
        description: `Columns detected: ${columns.map(col => `${col.name} (${col.type})`).join(', ')}`,
        importance: 'low'
      })
    }

    return newInsights.slice(0, 8) // Limit to 8 insights
  }, [generateTimeSeriesInsights])

  const processExcelFile = useCallback(async (file: File) => {
    setIsProcessing(true)
    setProcessingStep('Reading file...')
    
    try {
      const buffer = await file.arrayBuffer()
      setProcessingStep('Parsing Excel data...')
      
      const workbook = XLSX.read(buffer, { type: 'buffer' })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
      
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
        
        // Try to detect dates first
        const dateValues: Date[] = []
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
          const sortedDates = [...dateValues].sort((a, b) => a.getTime() - b.getTime())
          const frequency = analyzeDateFrequency(dateValues)
          
          const minDate = sortedDates[0]
          const maxDate = sortedDates[sortedDates.length - 1]
          const dateRange = minDate && maxDate ? 
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
            }
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
              sum: numericValues.reduce((a, b) => a + b, 0)
            }
          }
        }
        
        return {
          name: header,
          type,
          values,
          stats
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

  const generateChart = (column: DataColumn, chartType: 'bar' | 'pie' | 'line' = 'bar') => {
    if (column.type === 'date') {
      // Create time series chart for date columns
      if (!column.dateValues) return null
      
      const timeSeriesData = column.dateValues
        .map((date, index) => ({
          date: date.toLocaleDateString(),
          value: 1,
          originalDate: date
        }))
        .sort((a, b) => a.originalDate.getTime() - b.originalDate.getTime())
      
      // Group by time period if too many points
      let groupedData = timeSeriesData
      if (timeSeriesData.length > 20) {
        const groups = new Map<string, number>()
        timeSeriesData.forEach(item => {
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

    // Prepare data for charts
    const chartData = column.values.map((value, index) => ({
      name: `Item ${index + 1}`,
      value: Number(value)
    })).slice(0, 10) // Limit to first 10 items for readability

    if (chartType === 'bar') {
      return (
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" />
            <YAxis />
            <Bar dataKey="value" fill={CHART_COLORS[0]} />
          </BarChart>
        </ResponsiveContainer>
      )
    } else if (chartType === 'pie') {
      return (
        <ResponsiveContainer width="100%" height={200}>
          <RechartsPieChart>
            <Pie data={chartData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80}>
              {chartData.map((_, index) => (
                <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
              ))}
            </Pie>
          </RechartsPieChart>
        </ResponsiveContainer>
      )
    }

    return null
  }
  
  // Generate time series chart combining date and numeric columns
  const generateTimeSeriesChart = (dateColumn: DataColumn, numericColumn: DataColumn) => {
    if (!dateColumn.dateValues || dateColumn.values.length !== numericColumn.values.length) {
      return null
    }
    
    const timeSeriesData: TimeSeriesData[] = []
    for (let i = 0; i < dateColumn.dateValues.length; i++) {
      const date = dateColumn.dateValues[i]
      const value = Number(numericColumn.values[i])
      if (date && !isNaN(value)) {
        timeSeriesData.push({
          date: date.toLocaleDateString(),
          value,
          originalDate: date
        })
      }
    }
    
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

  const exportInfographic = useCallback(() => {
    // Create a simple HTML export
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Data Analysis Report - ${fileName}</title>
        <style>
          body { font-family: 'Inter', sans-serif; margin: 40px; background: #f8fafc; }
          .container { max-width: 800px; margin: 0 auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
          .header { text-align: center; margin-bottom: 40px; }
          .insight { padding: 20px; margin: 20px 0; background: #f1f5f9; border-radius: 8px; border-left: 4px solid #1e40af; }
          .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 30px 0; }
          .stat-card { padding: 20px; background: #1e40af; color: white; border-radius: 8px; text-align: center; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>Data Analysis Report</h1>
            <p>Generated from: ${fileName}</p>
          </div>
          ${insights.map(insight => `
            <div class="insight">
              <h3>${insight.title}</h3>
              <p>${insight.description}</p>
              ${insight.value ? `<strong>Value: ${insight.value}</strong>` : ''}
            </div>
          `).join('')}
        </div>
      </body>
      </html>
    `
    
    const blob = new Blob([htmlContent], { type: 'text/html' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${fileName.replace(/\.[^/.]+$/, '')}_analysis_report.html`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
    
    toast.success('Infographic exported successfully!')
  }, [fileName, insights])

  const clearData = useCallback(() => {
    setUploadedData([])
    setInsights([])
    setFileName('')
    toast.success('Data cleared successfully!')
  }, [setUploadedData, setInsights, setFileName])

  if (uploadedData.length === 0) {
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
      <div className="max-w-6xl mx-auto">
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
          <TabsList className="grid w-full grid-cols-4">
            <TabsTrigger value="insights" className="flex items-center gap-2">
              <TrendingUp className="w-4 h-4" />
              Insights
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
          </TabsList>

          <TabsContent value="insights" className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {insights.map((insight, index) => (
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

          <TabsContent value="timeseries" className="space-y-6">
            {(() => {
              const dateColumns = uploadedData.filter(col => col.type === 'date')
              const numericColumns = uploadedData.filter(col => col.type === 'number')
              
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
                            <span className="font-medium">{dateCol.dateValues?.length}</span>
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
              {uploadedData.filter(col => col.type === 'number').map((column, index) => (
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
              
              {uploadedData.filter(col => col.type === 'date').map((column, index) => (
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
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {uploadedData.map((column, index) => (
                <Card key={index}>
                  <CardHeader>
                    <CardTitle className="text-lg">{column.name}</CardTitle>
                    <CardDescription>
                      Type: {column.type} • {column.stats?.count} values
                    </CardDescription>
                  </CardHeader>
                  <CardContent>
                    {column.type === 'number' && column.stats && (
                      <div className="space-y-2">
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Min:</span>
                          <span className="font-medium">{(column.stats.min as number)?.toLocaleString()}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Max:</span>
                          <span className="font-medium">{(column.stats.max as number)?.toLocaleString()}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Average:</span>
                          <span className="font-medium">
                            {column.stats.avg?.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                          </span>
                        </div>
                        <Separator />
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Sum:</span>
                          <span className="font-medium">{column.stats.sum?.toLocaleString()}</span>
                        </div>
                      </div>
                    )}
                    {column.type === 'date' && column.stats && (
                      <div className="space-y-2">
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Earliest:</span>
                          <span className="font-medium">{(column.stats.min as Date)?.toLocaleDateString()}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Latest:</span>
                          <span className="font-medium">{(column.stats.max as Date)?.toLocaleDateString()}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Pattern:</span>
                          <Badge variant="outline">{column.stats.frequency}</Badge>
                        </div>
                        <Separator />
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Span:</span>
                          <span className="font-medium">{column.stats.dateRange}</span>
                        </div>
                        {column.stats.formatSummary && (
                          <>
                            <Separator />
                            <div className="space-y-1">
                              <span className="text-muted-foreground text-sm">Detected Formats:</span>
                              <div className="text-xs text-muted-foreground break-words">
                                {column.stats.formatSummary}
                              </div>
                            </div>
                          </>
                        )}
                      </div>
                    )}
                    {column.type === 'text' && (
                      <div className="text-muted-foreground">
                        Preview: {column.values.slice(0, 3).join(', ')}
                        {column.values.length > 3 && '...'}
                      </div>
                    )}
                  </CardContent>
                </Card>
              ))}
            </div>
          </TabsContent>
        </Tabs>
      </div>
    </div>
  )
}

export default App