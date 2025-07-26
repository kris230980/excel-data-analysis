import React, { useState, useCallback } from 'react'
import { useKV } from '@github/spark/hooks'
import * as XLSX from 'xlsx'
import { Upload, BarChart3, Download, FileSpreadsheet, TrendingUp, PieChart } from '@phosphor-icons/react'
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
  stats?: {
    min?: number
    max?: number
    avg?: number
    sum?: number
    count: number
  }
}

interface Insight {
  type: 'trend' | 'correlation' | 'summary' | 'outlier'
  title: string
  description: string
  value?: string | number
  importance: 'high' | 'medium' | 'low'
}

const CHART_COLORS = ['#1e40af', '#f97316', '#10b981', '#8b5cf6', '#ef4444', '#06b6d4']

function App() {
  const [uploadedData, setUploadedData] = useKV<DataColumn[]>('excel-data', [])
  const [insights, setInsights] = useKV<Insight[]>('data-insights', [])
  const [isProcessing, setIsProcessing] = useState(false)
  const [processingStep, setProcessingStep] = useState('')
  const [fileName, setFileName] = useKV<string>('file-name', '')

  const analyzeData = useCallback((columns: DataColumn[]): Insight[] => {
    const newInsights: Insight[] = []
    
    // Find numeric columns for analysis
    const numericColumns = columns.filter(col => col.type === 'number' && col.stats)
    
    if (numericColumns.length > 0) {
      // Summary insight
      const totalRows = columns[0]?.values.length || 0
      newInsights.push({
        type: 'summary',
        title: 'Dataset Overview',
        description: `Analyzed ${totalRows} rows across ${columns.length} columns with ${numericColumns.length} numeric metrics`,
        value: `${totalRows} records`,
        importance: 'high'
      })

      // Find highest value column
      const maxColumn = numericColumns.reduce((max, col) => 
        (col.stats!.max || 0) > (max.stats!.max || 0) ? col : max
      )
      
      newInsights.push({
        type: 'trend',
        title: 'Highest Values',
        description: `${maxColumn.name} shows the highest peak value in your dataset`,
        value: maxColumn.stats!.max?.toLocaleString(),
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
    }

    return newInsights.slice(0, 5) // Limit to 5 insights
  }, [])

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
        
        // Determine column type
        const isNumeric = values.every(val => !isNaN(Number(val)) && val !== '')
        const type = isNumeric ? 'number' : 'text'
        
        // Calculate statistics for numeric columns
        let stats: DataColumn['stats'] = { count: values.length }
        
        if (type === 'number') {
          const numericValues = values.map(val => Number(val))
          stats = {
            count: values.length,
            min: Math.min(...numericValues),
            max: Math.max(...numericValues),
            avg: numericValues.reduce((a, b) => a + b, 0) / numericValues.length,
            sum: numericValues.reduce((a, b) => a + b, 0)
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
      
      // Update state
      setUploadedData(columns)
      setInsights(generatedInsights)
      setFileName(file.name)
      
      toast.success('File processed successfully!')
      
    } catch (error) {
      toast.error(`Error processing file: ${error instanceof Error ? error.message : 'Unknown error'}`)
    } finally {
      setIsProcessing(false)
      setProcessingStep('')
    }
  }, [analyzeData, setUploadedData, setInsights, setFileName])

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
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="text-center p-4">
                      <Upload className="w-8 h-8 text-primary mx-auto mb-2" />
                      <h3 className="font-medium">1. Upload</h3>
                      <p className="text-sm text-muted-foreground">Upload your Excel or CSV file</p>
                    </div>
                    <div className="text-center p-4">
                      <TrendingUp className="w-8 h-8 text-primary mx-auto mb-2" />
                      <h3 className="font-medium">2. Analyze</h3>
                      <p className="text-sm text-muted-foreground">Get automated insights and statistics</p>
                    </div>
                    <div className="text-center p-4">
                      <Download className="w-8 h-8 text-primary mx-auto mb-2" />
                      <h3 className="font-medium">3. Export</h3>
                      <p className="text-sm text-muted-foreground">Download beautiful infographics</p>
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
          <TabsList className="grid w-full grid-cols-3">
            <TabsTrigger value="insights" className="flex items-center gap-2">
              <TrendingUp className="w-4 h-4" />
              Insights
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
                        <span className="font-medium">{column.stats?.max?.toLocaleString()}</span>
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
                          <span className="font-medium">{column.stats.min?.toLocaleString()}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-muted-foreground">Max:</span>
                          <span className="font-medium">{column.stats.max?.toLocaleString()}</span>
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