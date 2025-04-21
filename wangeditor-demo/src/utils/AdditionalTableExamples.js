/**
 * 新增的表格示例集合
 */

/**
 * 获取项目进度表格示例
 * @returns {String} HTML格式的表格示例
 */
export function getProjectProgressTableHTML() {
  return `
<h3 style="text-align:center;color:#1976d2;">2024年项目进度追踪表</h3>
<table border="1" style="width:100%;border-collapse:collapse;margin:20px 0;">
  <tr>
    <th rowspan="2" style="background-color:#e3f2fd;color:#0d47a1;padding:10px;text-align:center;">项目名称</th>
    <th rowspan="2" style="background-color:#e3f2fd;color:#0d47a1;padding:10px;text-align:center;">负责团队</th>
    <th colspan="4" style="background-color:#e3f2fd;color:#0d47a1;padding:10px;text-align:center;">各季度完成情况</th>
    <th rowspan="2" style="background-color:#e3f2fd;color:#0d47a1;padding:10px;text-align:center;">总体进度</th>
  </tr>
  <tr>
    <th style="background-color:#bbdefb;padding:8px;">Q1</th>
    <th style="background-color:#bbdefb;padding:8px;">Q2</th>
    <th style="background-color:#bbdefb;padding:8px;">Q3</th>
    <th style="background-color:#bbdefb;padding:8px;">Q4</th>
  </tr>
  <tr>
    <td style="font-weight:bold;">网站重构项目</td>
    <td style="text-align:center;">前端开发组</td>
    <td style="text-align:center;background-color:#c8e6c9;color:#2e7d32;">100%</td>
    <td style="text-align:center;background-color:#c8e6c9;color:#2e7d32;">100%</td>
    <td style="text-align:center;background-color:#ffeb3b;color:#f57f17;">85%</td>
    <td style="text-align:center;background-color:#ffcdd2;color:#c62828;">40%</td>
    <td style="text-align:center;font-weight:bold;color:#ff9800;">81.25%</td>
  </tr>
  <tr>
    <td style="font-weight:bold;">移动应用开发</td>
    <td style="text-align:center;">移动开发组</td>
    <td style="text-align:center;background-color:#c8e6c9;color:#2e7d32;">100%</td>
    <td style="text-align:center;background-color:#c8e6c9;color:#2e7d32;">90%</td>
    <td style="text-align:center;background-color:#ffeb3b;color:#f57f17;">65%</td>
    <td style="text-align:center;background-color:#ffcdd2;color:#c62828;">30%</td>
    <td style="text-align:center;font-weight:bold;color:#ff5722;">71.25%</td>
  </tr>
  <tr>
    <td style="font-weight:bold;">数据分析平台</td>
    <td style="text-align:center;">数据团队</td>
    <td style="text-align:center;background-color:#c8e6c9;color:#2e7d32;">100%</td>
    <td style="text-align:center;background-color:#c8e6c9;color:#2e7d32;">100%</td>
    <td style="text-align:center;background-color:#c8e6c9;color:#2e7d32;">100%</td>
    <td style="text-align:center;background-color:#ffeb3b;color:#f57f17;">75%</td>
    <td style="text-align:center;font-weight:bold;color:#4caf50;">93.75%</td>
  </tr>
  <tr>
    <td colspan="6" style="background-color:#f5f5f5;font-weight:bold;text-align:right;padding-right:10px;">整体项目完成率</td>
    <td style="background-color:#e8f5e9;color:#1b5e20;font-weight:bold;text-align:center;font-size:16px;">82.08%</td>
  </tr>
</table>
<p style="margin-top:20px;font-size:14px;color:#666;">
<strong>说明：</strong>项目进度表展示了多个项目在各季度的完成情况，使用颜色编码来直观展示进度状态。
</p>`;
}

/**
 * 获取人力资源表格示例
 * @returns {String} HTML格式的表格示例
 */
export function getHRDepartmentTableHTML() {
  return `
<h3 style="text-align:center;color:#ff5722;">人力资源部门组织架构表</h3>
<table border="1" style="width:100%;border-collapse:collapse;margin:20px 0;">
  <tr>
    <th style="background-color:#fbe9e7;color:#d84315;padding:12px;text-align:center;font-size:16px;">岗位</th>
    <th style="background-color:#fbe9e7;color:#d84315;padding:12px;text-align:center;font-size:16px;">员工姓名</th>
    <th style="background-color:#fbe9e7;color:#d84315;padding:12px;text-align:center;font-size:16px;">入职时间</th>
    <th style="background-color:#fbe9e7;color:#d84315;padding:12px;text-align:center;font-size:16px;">主要职责</th>
    <th style="background-color:#fbe9e7;color:#d84315;padding:12px;text-align:center;font-size:16px;">绩效评级</th>
  </tr>
  <tr>
    <td rowspan="2" style="font-weight:bold;vertical-align:middle;text-align:center;background-color:#fff3e0;">HR总监</td>
    <td style="text-align:center;">王小明</td>
    <td style="text-align:center;">2018-03-15</td>
    <td>
      <table border="1" style="width:100%;border-collapse:collapse;margin:5px 0;">
        <tr>
          <td style="background-color:#e8f5e9;padding:5px;">• 制定人力资源战略</td>
        </tr>
        <tr>
          <td style="background-color:#e8f5e9;padding:5px;">• 管理各职能模块负责人</td>
        </tr>
        <tr>
          <td style="background-color:#e8f5e9;padding:5px;">• 优化组织结构</td>
        </tr>
      </table>
    </td>
    <td style="text-align:center;color:#388e3c;font-weight:bold;">A+</td>
  </tr>
  <tr>
    <td style="text-align:center;font-style:italic;color:#9e9e9e;">(副总监)</td>
    <td style="text-align:center;font-style:italic;color:#9e9e9e;">招聘中</td>
    <td>-</td>
    <td style="text-align:center;color:#9e9e9e;">-</td>
  </tr>
  <tr>
    <td rowspan="3" style="font-weight:bold;vertical-align:middle;text-align:center;background-color:#e3f2fd;">招聘专员</td>
    <td style="text-align:center;">李梅</td>
    <td style="text-align:center;">2020-07-01</td>
    <td>• 处理日常招聘<br>• 维护人才库<br>• 组织面试</td>
    <td style="text-align:center;color:#1976d2;font-weight:bold;">A</td>
  </tr>
  <tr>
    <td style="text-align:center;">张强</td>
    <td style="text-align:center;">2021-09-15</td>
    <td>• 校园招聘<br>• 雇主品牌建设<br>• 实习生项目</td>
    <td style="text-align:center;color:#1976d2;font-weight:bold;">A-</td>
  </tr>
  <tr>
    <td style="text-align:center;">赵丽</td>
    <td style="text-align:center;">2023-03-20</td>
    <td>• 社会招聘<br>• 简历筛选<br>• 面试安排</td>
    <td style="text-align:center;color:#ff9800;font-weight:bold;">B+</td>
  </tr>
</table>
<p style="margin-top:20px;font-size:14px;color:#666;">
<strong>说明：</strong>人力资源部门表格展示了组织架构、人员配置、职责划分和绩效情况，包含嵌套表格显示详细职责。
</p>`;
}

/**
 * 获取财务报表示例
 * @returns {String} HTML格式的表格示例
 */
export function getFinancialReportTableHTML() {
  return `
<h3 style="text-align:center;color:#4caf50;">2024年第一季度财务报表摘要</h3>
<table border="1" style="width:100%;border-collapse:collapse;margin:20px 0;">
  <tr>
    <th colspan="2" style="background-color:#e8f5e9;color:#1b5e20;padding:12px;text-align:center;font-size:16px;">资产负债表</th>
    <th colspan="2" style="background-color:#e8f5e9;color:#1b5e20;padding:12px;text-align:center;font-size:16px;">利润表</th>
  </tr>
  <tr>
    <th style="background-color:#c8e6c9;padding:8px;">项目</th>
    <th style="background-color:#c8e6c9;padding:8px;">金额（万元）</th>
    <th style="background-color:#c8e6c9;padding:8px;">项目</th>
    <th style="background-color:#c8e6c9;padding:8px;">金额（万元）</th>
  </tr>
  <tr>
    <td style="font-weight:bold;">流动资产</td>
    <td style="text-align:right;padding-right:10px;">5,600.00</td>
    <td style="font-weight:bold;">营业收入</td>
    <td style="text-align:right;padding-right:10px;color:#2e7d32;font-weight:bold;">8,900.00</td>
  </tr>
  <tr>
    <td style="padding-left:20px;">- 货币资金</td>
    <td style="text-align:right;padding-right:10px;">2,800.00</td>
    <td style="padding-left:20px;">- 主营业务收入</td>
    <td style="text-align:right;padding-right:10px;">8,200.00</td>
  </tr>
  <tr>
    <td style="padding-left:20px;">- 应收账款</td>
    <td style="text-align:right;padding-right:10px;">1,500.00</td>
    <td style="padding-left:20px;">- 其他业务收入</td>
    <td style="text-align:right;padding-right:10px;">700.00</td>
  </tr>
  <tr>
    <td style="padding-left:20px;">- 存货</td>
    <td style="text-align:right;padding-right:10px;">1,300.00</td>
    <td style="font-weight:bold;">营业成本</td>
    <td style="text-align:right;padding-right:10px;color:#d32f2f;font-weight:bold;">6,300.00</td>
  </tr>
  <tr>
    <td style="font-weight:bold;">非流动资产</td>
    <td style="text-align:right;padding-right:10px;">7,200.00</td>
    <td style="padding-left:20px;">- 主营业务成本</td>
    <td style="text-align:right;padding-right:10px;">5,800.00</td>
  </tr>
  <tr>
    <td style="padding-left:20px;">- 固定资产</td>
    <td style="text-align:right;padding-right:10px;">5,000.00</td>
    <td style="padding-left:20px;">- 其他业务成本</td>
    <td style="text-align:right;padding-right:10px;">500.00</td>
  </tr>
  <tr>
    <td style="padding-left:20px;">- 无形资产</td>
    <td style="text-align:right;padding-right:10px;">2,200.00</td>
    <td style="font-weight:bold;">税金及附加</td>
    <td style="text-align:right;padding-right:10px;">180.00</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#fff9c4;">资产总计</td>
    <td style="text-align:right;padding-right:10px;background-color:#fff9c4;font-weight:bold;">12,800.00</td>
    <td style="font-weight:bold;">期间费用</td>
    <td style="text-align:right;padding-right:10px;">780.00</td>
  </tr>
  <tr>
    <td style="font-weight:bold;">流动负债</td>
    <td style="text-align:right;padding-right:10px;">3,200.00</td>
    <td style="padding-left:20px;">- 销售费用</td>
    <td style="text-align:right;padding-right:10px;">380.00</td>
  </tr>
  <tr>
    <td style="font-weight:bold;">非流动负债</td>
    <td style="text-align:right;padding-right:10px;">2,100.00</td>
    <td style="padding-left:20px;">- 管理费用</td>
    <td style="text-align:right;padding-right:10px;">300.00</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#ffcdd2;">负债合计</td>
    <td style="text-align:right;padding-right:10px;background-color:#ffcdd2;font-weight:bold;">5,300.00</td>
    <td style="padding-left:20px;">- 财务费用</td>
    <td style="text-align:right;padding-right:10px;">100.00</td>
  </tr>
  <tr>
    <td style="font-weight:bold;">所有者权益</td>
    <td style="text-align:right;padding-right:10px;">7,500.00</td>
    <td style="font-weight:bold;background-color:#c8e6c9;">营业利润</td>
    <td style="text-align:right;padding-right:10px;background-color:#c8e6c9;font-weight:bold;color:#388e3c;">1,640.00</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#fff9c4;">负债和所有者权益总计</td>
    <td style="text-align:right;padding-right:10px;background-color:#fff9c4;font-weight:bold;">12,800.00</td>
    <td style="font-weight:bold;background-color:#c8e6c9;">净利润</td>
    <td style="text-align:right;padding-right:10px;background-color:#c8e6c9;font-weight:bold;color:#388e3c;">1,230.00</td>
  </tr>
</table>
<p style="margin-top:20px;font-size:14px;color:#666;">
<strong>说明：</strong>财务报表展示了资产负债表和利润表的主要项目，使用颜色区分不同类型的科目，便于快速理解财务状况。
</p>`;
}

/**
 * 获取产品对比表格示例
 * @returns {String} HTML格式的表格示例
 */
export function getProductComparisonTableHTML() {
  return `
<h3 style="text-align:center;color:#673ab7;">智能手机产品线对比表</h3>
<table border="1" style="width:100%;border-collapse:collapse;margin:20px 0;">
  <tr>
    <th style="background-color:#ede7f6;color:#4527a0;padding:12px;text-align:center;">规格/型号</th>
    <th style="background-color:#ede7f6;color:#4527a0;padding:12px;text-align:center;">入门版</th>
    <th style="background-color:#ede7f6;color:#4527a0;padding:12px;text-align:center;">标准版</th>
    <th style="background-color:#ede7f6;color:#4527a0;padding:12px;text-align:center;">专业版</th>
    <th style="background-color:#ede7f6;color:#4527a0;padding:12px;text-align:center;">旗舰版</th>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#f5f5f5;">处理器</td>
    <td style="text-align:center;">Snapdragon 680</td>
    <td style="text-align:center;">Snapdragon 778G</td>
    <td style="text-align:center;">Snapdragon 888</td>
    <td style="text-align:center;color:#673ab7;font-weight:bold;">Snapdragon 8 Gen 2</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#f5f5f5;">运行内存</td>
    <td style="text-align:center;">4GB</td>
    <td style="text-align:center;">6GB</td>
    <td style="text-align:center;">8GB</td>
    <td style="text-align:center;color:#673ab7;font-weight:bold;">12GB</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#f5f5f5;">存储空间</td>
    <td style="text-align:center;">64GB</td>
    <td style="text-align:center;">128GB</td>
    <td style="text-align:center;">256GB</td>
    <td style="text-align:center;color:#673ab7;font-weight:bold;">512GB</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#f5f5f5;">屏幕</td>
    <td>
      <table border="1" style="width:100%;border-collapse:collapse;margin:0;">
        <tr>
          <td style="padding:5px;">• 6.5英寸 LCD</td>
        </tr>
        <tr>
          <td style="padding:5px;">• 60Hz刷新率</td>
        </tr>
        <tr>
          <td style="padding:5px;">• 1080P分辨率</td>
        </tr>
      </table>
    </td>
    <td>
      <table border="1" style="width:100%;border-collapse:collapse;margin:0;">
        <tr>
          <td style="padding:5px;">• 6.67英寸 AMOLED</td>
        </tr>
        <tr>
          <td style="padding:5px;">• 90Hz刷新率</td>
        </tr>
        <tr>
          <td style="padding:5px;">• 1080P分辨率</td>
        </tr>
      </table>
    </td>
    <td>
      <table border="1" style="width:100%;border-collapse:collapse;margin:0;">
        <tr>
          <td style="padding:5px;">• 6.8英寸 AMOLED</td>
        </tr>
        <tr>
          <td style="padding:5px;">• 120Hz刷新率</td>
        </tr>
        <tr>
          <td style="padding:5px;">• 2K分辨率</td>
        </tr>
      </table>
    </td>
    <td>
      <table border="1" style="width:100%;border-collapse:collapse;margin:0;">
        <tr>
          <td style="padding:5px;color:#673ab7;font-weight:bold;">• 6.8英寸 LTPO AMOLED</td>
        </tr>
        <tr>
          <td style="padding:5px;color:#673ab7;font-weight:bold;">• 144Hz自适应刷新率</td>
        </tr>
        <tr>
          <td style="padding:5px;color:#673ab7;font-weight:bold;">• 2K分辨率</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#f5f5f5;">相机系统</td>
    <td style="text-align:center;">主摄48MP<br>广角8MP</td>
    <td style="text-align:center;">主摄64MP<br>广角13MP<br>微距5MP</td>
    <td style="text-align:center;">主摄108MP<br>广角16MP<br>长焦12MP<br>微距5MP</td>
    <td style="text-align:center;color:#673ab7;font-weight:bold;">主摄200MP<br>广角32MP<br>长焦48MP<br>微距5MP<br>ToF传感器</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#f5f5f5;">电池容量</td>
    <td style="text-align:center;">4500mAh<br>18W快充</td>
    <td style="text-align:center;">5000mAh<br>33W快充</td>
    <td style="text-align:center;">5000mAh<br>67W快充</td>
    <td style="text-align:center;color:#673ab7;font-weight:bold;">5500mAh<br>120W超级快充<br>50W无线充电</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#f5f5f5;">特色功能</td>
    <td>• 指纹识别<br>• 面部解锁</td>
    <td>• 指纹识别<br>• 面部解锁<br>• NFC</td>
    <td>• 屏下指纹<br>• 面部解锁<br>• NFC<br>• 防水防尘</td>
    <td style="color:#673ab7;font-weight:bold;">• 屏下指纹<br>• 3D面部识别<br>• NFC<br>• IP68防水<br>• 无线充电<br>• 反向充电</td>
  </tr>
  <tr>
    <td style="font-weight:bold;background-color:#e8eaf6;">价格</td>
    <td style="text-align:center;font-weight:bold;color:#4caf50;">￥1,999</td>
    <td style="text-align:center;font-weight:bold;color:#ff9800;">￥2,999</td>
    <td style="text-align:center;font-weight:bold;color:#f44336;">￥4,999</td>
    <td style="text-align:center;font-weight:bold;color:#e91e63;font-size:18px;">￥7,999</td>
  </tr>
</table>
<p style="margin-top:20px;font-size:14px;color:#666;">
<strong>说明：</strong>产品对比表使用嵌套表格展示详细规格信息，通过颜色和样式突出显示各型号特点和价格差异。
</p>`;
}
