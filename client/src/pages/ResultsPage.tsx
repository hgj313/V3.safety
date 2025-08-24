import React, { useState, useEffect, useRef } from 'react';
import { 
  Alert, 
  Tabs, 
  Typography, 
  Button,
  Space,
  message,
  Skeleton,
  Row,
  Col,
  Card
} from 'antd';
import { 
  FileExcelOutlined, 
  FilePdfOutlined
} from '@ant-design/icons';
import { useOptimizationContext } from '../contexts/OptimizationContext';
import { useOptimizationResults } from '../hooks/useOptimizationResults';
import ResultsOverview from '../components/results/ResultsOverview';
import CuttingPlansTable from '../components/results/CuttingPlansTable';
import RequirementsValidation from '../components/results/RequirementsValidation';
import ProcurementList from '../components/results/ProcurementList';

const { Title } = Typography;
const { TabPane } = Tabs;

const ResultsPage: React.FC = () => {
  const { currentOptimization, designSteels, moduleSteels } = useOptimizationContext();
  const results = currentOptimization?.results;
  const [activeTab, setActiveTab] = useState('overview');
  const [exporting, setExporting] = useState(false);
  const hasWarnedRef = useRef(false);

  // 使用统一的数据处理Hook - 这是解决错误引用值问题的核心
  const processedResults = useOptimizationResults(results, designSteels, moduleSteels);

  // 自动化调试输出和错误检测
  useEffect(() => {
    console.log('【ResultsPage V3】自动排查开始');
    console.log('【数据源】results:', results);
    console.log('【数据源】designSteels:', designSteels);
    console.log('【数据源】moduleSteels:', moduleSteels);
    console.log('【处理结果】processedResults:', processedResults);
    
    // 数据一致性验证
    if (results?.solutions) {
      console.log('【一致性检查】后端统计 vs 前端处理:');
      console.log('- 后端totalModuleUsed:', results.totalModuleUsed);
      console.log('- 前端totalModuleCount:', processedResults.totalStats.totalModuleCount);
      console.log('- 后端totalMaterial:', results.totalMaterial);
      console.log('- 前端totalModuleLength:', processedResults.totalStats.totalModuleLength);
      console.log('- 后端totalWaste:', results.totalWaste);
      console.log('- 前端totalWaste:', processedResults.totalStats.totalWaste);
      
      // 检查数据一致性
      if (results.totalModuleUsed !== processedResults.totalStats.totalModuleCount) {
        console.warn('⚠️ 数据不一致：模数钢材用量');
      }
      if (results.totalMaterial !== processedResults.totalStats.totalModuleLength) {
        console.warn('⚠️ 数据不一致：模数钢材总长度');
      }
      if (results.totalWaste !== processedResults.totalStats.totalWaste) {
        console.warn('⚠️ 数据不一致：废料统计');
      }
    }

    // 错误提示
    if (processedResults.hasDataError) {
      message.error(`数据异常：${processedResults.errorMessage}`);
    } else if (!processedResults.isAllRequirementsSatisfied) {
      if (!hasWarnedRef.current) {
        message.warning('部分需求未满足，请检查优化配置');
        hasWarnedRef.current = true;
      }
    } else {
      hasWarnedRef.current = false;
      console.log('✅ 数据验证通过，所有需求已满足');
    }

    if (currentOptimization?.error) {
      message.error('后端错误：' + currentOptimization.error);
    }
    
    console.log('【ResultsPage V3】自动排查完成');
  }, [results, designSteels, moduleSteels, processedResults, currentOptimization]);

  // 错误状态处理
  if (processedResults.hasDataError) {
    return (
      <div style={{ padding: '20px' }}>
        <Alert
          message="数据异常"
          description={processedResults.errorMessage || '优化结果数据异常，请重新执行优化'}
          type="error"
          showIcon
          action={
            <div style={{ display: 'flex', justifyContent: 'center' }}>
              <Button size="small" onClick={() => window.location.reload()}>
                刷新页面
              </Button>
            </div>
          }
        />
      </div>
    );
  }

  // 空数据或加载中状态处理
  if (!results || !results.solutions) {
    return (
      <div style={{ padding: '24px' }}>
        <Title level={2}>
          <Skeleton.Input style={{ width: 200, marginBottom: 24 }} active />
        </Title>
        
        {/* High-fidelity Skeleton for Tabs and content */}
        <div>
          {/* Skeleton for Tab headers */}
          <div style={{ display: 'flex', borderBottom: '1px solid #f0f0f0', marginBottom: 24 }}>
            <Skeleton.Button active style={{ width: 100, marginRight: 16 }} />
            <Skeleton.Button active style={{ width: 100, marginRight: 16 }} />
            <Skeleton.Button active style={{ width: 100, marginRight: 16 }} />
            <Skeleton.Button active style={{ width: 100 }} />
          </div>
          
          {/* Skeleton for Tab content (mimicking ResultsOverview) */}
          <div>
            <Row gutter={[16, 16]} style={{ marginBottom: 24 }}>
              <Col span={8}><Card><Skeleton active paragraph={{ rows: 2 }} /></Card></Col>
              <Col span={8}><Card><Skeleton active paragraph={{ rows: 2 }} /></Card></Col>
              <Col span={8}><Card><Skeleton active paragraph={{ rows: 2 }} /></Card></Col>
            </Row>
            <Row>
              <Col span={24}>
                <Card>
                  <Skeleton.Node active style={{ height: 300, width: '100%' }}>
                    <div />
                  </Skeleton.Node>
                </Card>
              </Col>
            </Row>
          </div>
        </div>
      </div>
    );
  }

  // 导出Excel功能
  const handleExportExcel = async () => {
    try {
      setExporting(true);
      
      // 验证数据 - 即使没有优化结果，只要有采购清单数据也可以导出
      if (!results || (!results.solutions && !processedResults.moduleUsageStats.sortedStats.length)) {
        throw new Error('没有可用的采购数据，请确保已完成优化或已有采购清单');
      }

      // 直接使用前端的moduleUsageStats数据，确保有采购清单
      const exportResults = {
        ...results,
        solutions: results?.solutions || [],
        // 使用前端计算好的moduleUsageStats
        moduleUsageStats: processedResults.moduleUsageStats.sortedStats.map(item => ({
          specification: item.specification,
          length: typeof item.length === 'number' ? item.length : parseInt(String(item.length), 10) || 0,
          totalUsed: item.count,
          averageUtilization: 0.95, // 默认利用率
          totalLength: item.totalLength
        })),
        summary: results?.summary || {},
        optimizationDetails: results?.optimizationDetails || {},
        // 添加前端统计数据
        frontendStats: {
          totalModuleCount: processedResults.totalStats.totalModuleCount,
          totalModuleLength: processedResults.totalStats.totalModuleLength,
          grandTotal: processedResults.moduleUsageStats.grandTotal
        }
      };

      console.log('📊 准备导出Excel数据:', {
        hasResults: !!results,
        hasSolutions: !!results?.solutions,
        hasModuleUsageStats: processedResults.moduleUsageStats.sortedStats.length > 0,
        moduleUsageStatsCount: processedResults.moduleUsageStats.sortedStats.length,
        grandTotalCount: processedResults.moduleUsageStats.grandTotal.count,
        exportOptions: {
          format: 'excel',
          includeCharts: false,
          includeDetails: true,
          includeLossRateBreakdown: true,
          customTitle: `钢材优化报告_${new Date().toISOString().slice(0, 10)}`
        }
      });
      
      // 确保发送正确的数据结构
      const exportData = {
        results: exportResults,
        exportOptions: {
          format: 'excel',
          includeCharts: false,
          includeDetails: true,
          includeLossRateBreakdown: true,
          customTitle: `钢材优化报告_${new Date().toISOString().slice(0, 10)}`,
          // 添加前端数据作为备选
          useFrontendData: true
        }
      };
      
      // 添加测试数据验证
      console.log('📤 发送数据验证:', {
        resultsExists: !!results,
        hasModuleUsageStats: exportResults.moduleUsageStats?.length > 0,
        moduleUsageStatsCount: exportResults.moduleUsageStats?.length || 0,
        grandTotalCount: processedResults.moduleUsageStats.grandTotal.count,
        exportDataKeys: Object.keys(exportData),
        bodySize: JSON.stringify(exportData).length
      });

      // 检测环境并选择正确的端点
      const isNetlify = window.location.hostname.includes('netlify.app') || 
                       window.location.hostname.includes('vercel.app') ||
                       !window.location.hostname.includes('localhost');
      
      const endpoint = isNetlify ? '/.netlify/functions/export-excel' : '/api/export/excel';

      const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(exportData),
      });

      if (!response.ok) {
        const errorText = await response.text();
        let errorMessage = 'Excel导出失败';
        try {
          const errorData = JSON.parse(errorText);
          errorMessage = errorData.error || errorMessage;
        } catch {
          errorMessage = errorText || errorMessage;
        }
        throw new Error(errorMessage);
      }

      // 获取文件名
      const contentDisposition = response.headers.get('Content-Disposition');
      let filename = '钢材优化报告.xlsx';
      if (contentDisposition) {
        const filenameMatch = contentDisposition.match(/filename[^;=]*=((['"]).*?\2|[^;\n]*)/);
        if (filenameMatch) {
          filename = decodeURIComponent(filenameMatch[1].replace(/['"]/g, ''));
        }
      }

      // 下载文件
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      message.success('Excel报告导出成功！');
    } catch (error) {
      console.error('Excel导出失败:', error);
      message.error(`Excel导出失败: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setExporting(false);
    }
  };

  // 导出PDF功能
  const handleExportPDF = async () => {
    try {
      setExporting(true);
      message.loading({ content: '正在生成报告...', key: 'export' });
      
      const exportData = {
        results: results,
        exportOptions: {
          format: 'pdf',
          includeDetails: true,
          includePurchaseList: true,
          includeDesignSteels: true
        },
        designSteels: designSteels
      };

      // 检测环境并选择正确的端点
      const isNetlify = window.location.hostname.includes('netlify.app') || 
                       window.location.hostname.includes('vercel.app') ||
                       !window.location.hostname.includes('localhost');
      
      const endpoint = isNetlify ? '/.netlify/functions/export-pdf' : '/api/export/pdf';

      const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(exportData),
      });

      if (!response.ok) {
        const errorText = await response.text();
        let errorMessage = 'PDF导出失败';
        try {
          const errorData = JSON.parse(errorText);
          errorMessage = errorData.error || errorMessage;
        } catch {
          errorMessage = errorText || errorMessage;
        }
        throw new Error(errorMessage);
      }

      // 直接下载HTML文件作为PDF替代方案
      const htmlContent = await response.text();
      const blob = new Blob([htmlContent], { type: 'text/html' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = `钢材优化报告_${new Date().toISOString().split('T')[0]}.html`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      message.success({ content: '报告已成功下载！请在浏览器中打开并打印为PDF。', key: 'export', duration: 5 });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      console.error('报告导出失败:', errorMessage);
      message.error({ content: `导出失败: ${errorMessage}`, key: 'export', duration: 3 });
    } finally {
      setExporting(false);
    }
  };

  console.log('🔍 ResultsPage V3渲染信息:', {
    resultsKeys: Object.keys(results.solutions),
    designSteelsCount: designSteels?.length || 0,
    moduleSteelsCount: moduleSteels?.length || 0,
    processedStats: processedResults.totalStats,
    hasDataError: processedResults.hasDataError
  });

  return (
    <div style={{ padding: '24px' }}>
      <Title level={2}>化优化结果分析</Title>

      {/* 顶部需求满足提示条已由父级统一显示，避免重复 */}
      
      <Tabs activeKey={activeTab} onChange={setActiveTab}>
        <TabPane tab="概览图表" key="overview">
          <ResultsOverview
            totalStats={processedResults.totalStats}
            chartData={processedResults.chartData}
            isAllRequirementsSatisfied={processedResults.isAllRequirementsSatisfied}
          />
        </TabPane>

        <TabPane tab="切割方案" key="cutting">
          <CuttingPlansTable
            regroupedResults={processedResults.regroupedResults}
            designIdToDisplayIdMap={processedResults.designIdToDisplayIdMap}
          />
        </TabPane>

        <TabPane tab="需求验证" key="requirements">
          <RequirementsValidation
            requirementValidation={processedResults.requirementValidation}
            isAllRequirementsSatisfied={processedResults.isAllRequirementsSatisfied}
          />
        </TabPane>

        <TabPane tab="采购清单" key="procurement">
          <ProcurementList
            moduleUsageStats={processedResults.moduleUsageStats}
          />
        </TabPane>
      </Tabs>

      {/* 导出功能 */}
      <div style={{ marginTop: 24, textAlign: 'center' }}>
        <Space>
          <Button 
            type="primary" 
            icon={<FileExcelOutlined />} 
            size="large"
            loading={exporting}
            onClick={handleExportExcel}
          >
            导出采购清单(Excel)
          </Button>
          <Button 
            icon={<FilePdfOutlined />} 
            size="large"
            loading={exporting}
            onClick={handleExportPDF}
          >
            导出完整报告(PDF)
          </Button>
        </Space>
      </div>
    </div>
  );
};

export default ResultsPage;