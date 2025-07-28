# Tools
For working application

# 1. BOM Compare.bas
  The macro processing be used for compare the standard 2 BOMs download for The TEAMCENTER.
  
  # how to use ?
    1. New create a Excel file, ensure the default sheet name is "Sheet1"
    2. Add 2 new sheets and the name shoud be "Sheet2" and "Sheet3"
    3. Copy all content of BOM 1 and BOM 2 with format into the new created Excel "Sheet1" and "Sheet2"
    4. Press the "Alt+F11" open the VBA edit page
    5. Select the VBAProject (Excal file name) then click the right key of mouse "import file" select the BOM Compare.bas then close the page
    6. Back to Excel file press "Alt+f8" run this "CompareSheets" macro processing
    
   # Key Components Explained:
   
     1. 使用说明：
          比较Sheet1和Sheet2的E列
          匹配时比较C列
          不匹配时导出C/D/E列到Sheet3
          Sheet1的F列显示结果状态
          Sheet3自动清空并添加表头
          自动调整列宽
     2. 使用说明：
          三表比较逻辑：
          在Sheet1中遍历每一行
          在Sheet2中查找匹配的E列值
          在Sheet3中查找相同的E列值
          比较所有三表中对应行的C列值
          结果输出到Sheet4：
          A列：数据来源（Sheet1行号）
          B-D列：Sheet1的C/D/E值
          E-G列：Sheet2的C/D/E值
          H-J列：Sheet3的C/D/E值
          K列：C列是否全部相等（布尔值）
          L列：状态文本（"All Equal"或"Not Equal"）
          红色标记功能：
          在比较前重置所有源工作表的字体颜色为黑色
          当找到三表匹配的行时，将该行所有单元格字体设为红色
          使用Rows(i).Font.Color确保整行标记
          数据处理优化：
          限制处理范围在2-300行
          空行自动跳过
          输出表自动调整列宽
          布尔值列特殊格式化
          使用说明：
          准备工作：
          确保工作簿包含Sheet1、Sheet2、Sheet3和Sheet4
          数据应从第2行开始（第1行为标题）
          运行宏后：
          Sheet1、Sheet2、Sheet3中匹配的行会变为红色字体
          Sheet4包含所有三表匹配记录的详细比较结果
          消息框显示处理统计信息
          结果解读：
          完全匹配：三表E列相同且C列相同 → "All Equal"
          部分匹配：三表E列相同但C列不同 → "Not Equal"
          无三表匹配：不会出现在Sheet4中
          注意事项：
          性能考虑：
          三重嵌套循环处理300³ = 27,000,000种可能组合
          对于更大数据集，建议使用字典优化
          添加Application.ScreenUpdating = False加速执行
          自定义选项：
          修改行范围：更改For i = 2 To 300中的300值
          更改高亮颜色：替换vbRed为其他颜色常量
          调整输出列：修改Sheet4的表头数组和写入位置
          错误处理：
          添加工作表存在性检查
          处理空工作表情况
          此解决方案提供了全面的三表比对功能，直观的视觉标记（红色字体），以及清晰的比对结果输出，便于分析数据一致性。
