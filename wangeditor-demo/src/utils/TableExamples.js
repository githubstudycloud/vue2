/**
 * 表格示例数据
 */

/**
 * 获取HTML格式的嵌套表格示例
 * @returns {String} HTML格式的表格示例
 */
export function getNestedTableHTML() {
  return `
<h3 style="text-align:center;color:#2196f3;">综合数据报表示例</h3>
<table border="1" style="width:100%;border-collapse:collapse;margin:20px 0;">
  <!-- 多层表头 -->
  <tr>
    <th colspan="2" style="background-color:#f5f5f5;color:#333;padding:10px;text-align:center;font-size:18px;">部门信息</th>
    <th colspan="4" style="background-color:#f5f5f5;color:#333;padding:10px;text-align:center;font-size:18px;">项目详情</th>
  </tr>
  <tr>
    <th style="background-color:#e3f2fd;color:#1976d2;padding:8px;">部门名称</th>
    <th style="background-color:#e3f2fd;color:#1976d2;padding:8px;">负责人</th>
    <th style="background-color:#e3f2fd;color:#1976d2;padding:8px;">项目类型</th>
    <th style="background-color:#e3f2fd;color:#1976d2;padding:8px;">完成进度</th>
    <th style="background-color:#e3f2fd;color:#1976d2;padding:8px;">投入资源</th>
    <th style="background-color:#e3f2fd;color:#1976d2;padding:8px;">详细数据</th>
  </tr>
  
  <!-- 研发部数据 -->
  <tr>
    <td rowspan="3" style="color:#1e88e5;font-size:16px;font-weight:bold;vertical-align:middle;text-align:center;">研发部</td>
    <td rowspan="3" style="vertical-align:middle;text-align:center;">张三</td>
    <td style="background-color:#e8f5e9;color:#2e7d32;">前端开发</td>
    <td style="text-align:center;">
      <span style="color:#4caf50;font-weight:bold;">85%</span>
    </td>
    <td style="text-align:right;padding-right:10px;">
      <span style="font-family:Arial;color:#ff5722;">￥500,000</span>
    </td>
    <td>
      <!-- 嵌套表格 -->
      <table border="1" style="width:100%;border-collapse:collapse;margin:5px 0;">
        <tr>
          <th style="background-color:#f1f8e9;padding:5px;">人员</th>
          <th style="background-color:#f1f8e9;padding:5px;">进度</th>
          <th style="background-color:#f1f8e9;padding:5px;">评级</th>
        </tr>
        <tr>
          <td style="text-align:center;">5人</td>
          <td style="text-align:center;color:#4caf50;">按计划进行</td>
          <td style="text-align:center;color:#ff9800;">★★★★☆</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="background-color:#fff3e0;color:#ef6c00;">后端开发</td>
    <td style="text-align:center;">
      <span style="color:#ff9800;font-weight:bold;">72%</span>
    </td>
    <td style="text-align:right;padding-right:10px;">
      <span style="font-family:Arial;color:#ff5722;">￥800,000</span>
    </td>
    <td>
      <!-- 嵌套表格 -->
      <table border="1" style="width:100%;border-collapse:collapse;margin:5px 0;">
        <tr>
          <th style="background-color:#fff3e0;padding:5px;">人员</th>
          <th style="background-color:#fff3e0;padding:5px;">进度</th>
          <th style="background-color:#fff3e0;padding:5px;">评级</th>
        </tr>
        <tr>
          <td style="text-align:center;">7人</td>
          <td style="text-align:center;color:#ff9800;">延期风险</td>
          <td style="text-align:center;color:#f44336;">★★★☆☆</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="background-color:#e1f5fe;color:#0288d1;">数据分析</td>
    <td style="text-align:center;">
      <span style="color:#2196f3;font-weight:bold;">65%</span>
    </td>
    <td style="text-align:right;padding-right:10px;">
      <span style="font-family:Arial;color:#ff5722;">￥300,000</span>
    </td>
    <td>
      <!-- 嵌套表格 -->
      <table border="1" style="width:100%;border-collapse:collapse;margin:5px 0;">
        <tr>
          <th style="background-color:#e1f5fe;padding:5px;">人员</th>
          <th style="background-color:#e1f5fe;padding:5px;">进度</th>
          <th style="background-color:#e1f5fe;padding:5px;">评级</th>
        </tr>
        <tr>
          <td style="text-align:center;">3人</td>
          <td style="text-align:center;color:#2196f3;">正常</td>
          <td style="text-align:center;color:#607d8b;">★★★☆☆</td>
        </tr>
      </table>
    </td>
  </tr>
  
  <!-- 市场部数据 -->
  <tr>
    <td rowspan="2" style="color:#e91e63;font-size:16px;font-weight:bold;vertical-align:middle;text-align:center;">市场部</td>
    <td rowspan="2" style="vertical-align:middle;text-align:center;">李四</td>
    <td style="background-color:#fce4ec;color:#c2185b;">线上营销</td>
    <td style="text-align:center;">
      <span style="color:#4caf50;font-weight:bold;">110%</span>
    </td>
    <td style="text-align:right;padding-right:10px;">
      <span style="font-family:Arial;color:#ff5722;">￥1,200,000</span>
    </td>
    <td>
      <!-- 嵌套表格 -->
      <table border="1" style="width:100%;border-collapse:collapse;margin:5px 0;">
        <tr>
          <th colspan="2" style="background-color:#fce4ec;padding:5px;text-align:center;">KPI达成情况</th>
        </tr>
        <tr>
          <td style="padding:5px;">ROI</td>
          <td style="text-align:center;color:#4caf50;font-weight:bold;">210%</td>
        </tr>
        <tr>
          <td style="padding:5px;">新客户</td>
          <td style="text-align:center;color:#e91e63;font-weight:bold;">+1200</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="background-color:#f3e5f5;color:#7b1fa2;">线下活动</td>
    <td style="text-align:center;">
      <span style="color:#9c27b0;font-weight:bold;">95%</span>
    </td>
    <td style="text-align:right;padding-right:10px;">
      <span style="font-family:Arial;color:#ff5722;">￥850,000</span>
    </td>
    <td>
      <!-- 嵌套表格 -->
      <table border="1" style="width:100%;border-collapse:collapse;margin:5px 0;">
        <tr>
          <th colspan="2" style="background-color:#f3e5f5;padding:5px;text-align:center;">活动效果</th>
        </tr>
        <tr>
          <td style="padding:5px;">参与人数</td>
          <td style="text-align:center;color:#9c27b0;font-weight:bold;">2500+</td>
        </tr>
        <tr>
          <td style="padding:5px;">转化率</td>
          <td style="text-align:center;color:#7b1fa2;font-weight:bold;">18%</td>
        </tr>
      </table>
    </td>
  </tr>
  
  <!-- 财务汇总 -->
  <tr>
    <td colspan="3" style="background-color:#f5f5f5;font-weight:bold;text-align:right;padding-right:20px;">总计</td>
    <td style="background-color:#ffebee;color:#c62828;font-weight:bold;text-align:center;font-size:16px;">87.4%</td>
    <td style="background-color:#ffebee;color:#c62828;font-weight:bold;text-align:right;padding-right:10px;font-size:16px;">￥3,650,000</td>
    <td style="background-color:#f5f5f5;">
      <!-- 汇总表格 -->
      <table border="1" style="width:100%;border-collapse:collapse;margin:5px 0;">
        <tr>
          <th style="background-color:#cfd8dc;padding:5px;">指标</th>
          <th style="background-color:#cfd8dc;padding:5px;">状态</th>
        </tr>
        <tr>
          <td style="padding:5px;">整体进度</td>
          <td style="text-align:center;color:#4caf50;font-weight:bold;">正常</td>
        </tr>
        <tr>
          <td style="padding:5px;">风险等级</td>
          <td style="text-align:center;color:#ff9800;font-weight:bold;">中等</td>
        </tr>
        <tr>
          <td style="padding:5px;">总体ROI</td>
          <td style="text-align:center;color:#2196f3;font-weight:bold;">145%</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<p style="margin-top:20px;font-size:14px;color:#666;">
<strong>说明：</strong>本表格展示了多层表头、合并单元格、嵌套表格、各种文字样式（字体、颜色、大小、对齐方式）等特性。
您可以通过编辑器工具栏修改内容，并使用"导出为Excel"功能将表格导出。
导出功能会尽可能保留原有的格式和样式。
</p>`;
}

/**
 * 获取带有复杂样式的表格示例
 * @returns {String} HTML格式的表格示例
 */
export function getStyledTableHTML() {
  return `
<table border="1" style="width:100%;border-collapse:collapse;margin:20px 0;">
  <tr>
    <th colspan="4" style="background-color:#3f51b5;color:#fff;font-size:20px;text-align:center;padding:15px;">
      产品销售季度报表
    </th>
  </tr>
  <tr>
    <th style="background-color:#c5cae9;width:25%;">产品类别</th>
    <th style="background-color:#c5cae9;width:25%;">Q1销量</th>
    <th style="background-color:#c5cae9;width:25%;">Q2销量</th>
    <th style="background-color:#c5cae9;width:25%;">增长率</th>
  </tr>
  <tr>
    <td style="color:#d32f2f;font-weight:bold;">电子产品</td>
    <td style="text-align:right;">15,234</td>
    <td style="text-align:right;">18,456</td>
    <td style="text-align:center;color:#4caf50;font-weight:bold;">+21.1%</td>
  </tr>
  <tr>
    <td style="color:#1976d2;font-weight:bold;">服装配饰</td>
    <td style="text-align:right;">8,721</td>
    <td style="text-align:right;">9,015</td>
    <td style="text-align:center;color:#ff9800;font-weight:bold;">+3.4%</td>
  </tr>
  <tr>
    <td style="color:#7b1fa2;font-weight:bold;">食品饮料</td>
    <td style="text-align:right;">22,190</td>
    <td style="text-align:right;">20,458</td>
    <td style="text-align:center;color:#f44336;font-weight:bold;">-7.8%</td>
  </tr>
  <tr>
    <td colspan="3" style="background-color:#f5f5f5;text-align:right;font-weight:bold;padding-right:10px;">总计</td>
    <td style="background-color:#ffeb3b;color:#000;text-align:center;font-weight:bold;">+5.6%</td>
  </tr>
</table>`;
}