class Sheet

  attr_reader :name

  def initialize(sheet)
    @sheet = sheet
    @name  = sheet.name
    @font  = "Meiryo UI"
    @size  = 10

    @sheet.Cells.Font.Name = @font
    @sheet.Cells.Font.Size = @size
  end

  def acol(value)
    # return column alphabet name of value on first row
    values = @sheet.UsedRange.Rows(1).Value.first
    tgtcol = values.zip(1..,'A'..).assoc(value) # [["Key",1,"A"],...
    if tgtcol
      tgtcol.last
    else
      warn "#{self.class.name}##{__method__} #{value} is not found in #{@sheet.name}!"
      nil
    end
  end

  def ncol(value)
    # return column number of value on first row
    values = @sheet.UsedRange.Rows(1).Value.first
    tgtcol = values.zip(1..,'A'..).assoc(value) # [["Key",1,"A"],...
    if tgtcol
      tgtcol[1]
    else
      warn "#{self.class.name}##{__method__} #{value} is not found in #{@sheet.name}!"
      nil
    end
  end

  ## Excel column name <-> number
=begin
  #http://d.hatena.ne.jp/keyesberry/touch/20111229/p1
  #http://qiita.com/akihyro/items/432f63ad9dc90f415e2d
  def alpha2num(alphabets)
    [*'A'..alphabets].size
  end

  def num2alpha(number)
    alpha = 'A'
    (number-1).times {alpha.succ!}
    alpha
  end

  alphabets = %w(A B Z AA AB AZ BB AAA IV ZZZ XFD)
  numbers = alphabets.map { |alpha| alpha2num alpha }
  # => [1, 2, 26, 27, 28, 52, 54, 703, 256, 18278, 16384]
  numbers.map { |num| num2alpha num }
  # => ["A", "B", "Z", "AA", "AB", "AZ", "BB", "AAA", "IV", "ZZZ", "XFD"]
=end

  def num2alpha(num)
    (1..num).zip('A'..).last.last
  end

  def rgb(red,green,blue)
    # http://blade.nagaokaut.ac.jp/cgi-bin/scat.rb/ruby/ruby-list/50163
    # http://www.relief.jp/itnote/archives/000482.php
    # VBAの場合は rrggbb の並びではない
    red | (green << 8) | (blue << 16)
  end

  def make_graph(range_obj,charttype,
                 left: 400,top: 40,height: 428,width: 907,
                 title: nil,
                 xaxistitle:  nil, yaxistitle: nil,
                 reverseplotorder_value: nil,   # X Axis ExcelConst::XlValue
                 reverseplotorder_category: nil # Y Axis ExcelConst::XlCategory
                )
    # height: 428, width: 907 to fit Power Point slide area height: 15.1cm,width 32.31cm
    # 17.64cm 500pt -> 1cm = 28.354pt

    range_obj.Select # needed as for chart SetSourceData

    shapes = @sheet.Shapes # 指定されたシートのすべての**Shape** オブジェクトのコレクションです。
    # AddChart2 ドキュメントにグラフを追加します。 グラフを表す**Shape** オブジェクトを返し、指定されたコレクションに追加します。
    # https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addchart2
    # Style Use "-1" to get the default style for the chart type specified in XlChartType.
    graph_shape = shapes.AddChart2(Style: -1,XlChartType: charttype,Left: left,Top: top,Width: width,Height: height)

    # reverseplotorder
    # refer https://www.relief.jp/docs/excel-vba-chart-reverse-plot-order.html
    # Y Axis is xlcategory, X Axis is xlValue
    graph_shape.Chart.Axes(Type: ExcelConst::XlCategory).ReversePlotOrder = true if reverseplotorder_category

    graph_shape.Chart.Axes(Type: ExcelConst::XlValue).ReversePlotOrder    = true if reverseplotorder_value

    ## title
    graph_shape.Chart.ChartTitle.Text = title if title
    graph_shape.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = @font
    # graph_shape.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14 # 14 default size

    ## axis title
    # x axis
    if xaxistitle # Primary
      graph_shape.Chart.Axes(Type: ExcelConst::XlCategory,AxisGroup: ExcelConst::XlPrimary).HasTitle = true
      graph_shape.Chart.Axes(Type: ExcelConst::XlCategory,AxisGroup: ExcelConst::XlPrimary).AxisTitle.Characters.Text = xaxistitle
    end

    # y axis
    if yaxistitle # Primary
      graph_shape.Chart.Axes(Type: ExcelConst::XlValue   ,AxisGroup: ExcelConst::XlPrimary).HasTitle = true
      graph_shape.Chart.Axes(Type: ExcelConst::XlValue   ,AxisGroup: ExcelConst::XlPrimary).AxisTitle.Characters.Text = yaxistitle
    end

    graph_shape # return graph_shape object for subsequent processing
  end

  def put_values_in_row(values,start_range_name = "A1")
    @sheet.Range(start_range_name).Resize(1,values.size).Value = values
  end

  def put_values_in_column(values,start_range_name = "A1")
    # Note
    # Since in case "= values", FIRST element of values is set in rows
    # need to "= values.zip(0..)", 0: dummy data
    @sheet.Range(start_range_name).Resize(values.size,1).Value = values.zip(0..)
  end

  def put_values_in_rowcolumn(values,start_range_name = "A1")
    # values should be two dimensional array
    row_size = values.size
    col_size = values.first.size
    @sheet.Range(start_range_name).Resize(row_size,col_size).Value = values
  end

  def get_values_in_row(row,range_type = nil)
    # ranget_type: for future use
    # 列全体
    # UsedRange 1列目 -UsedRange最終列 現時点ではこれを対応
    # UsedRange 指定列-UsedRange最終列
    # ...
    @sheet.UsedRange.Rows(row).Value.first
  end

  def get_values_in_colum(column_name_or_number,range_type = nil)
    # ranget_type: for future use
    # 行全体
    # UsedRange 1行目-UsedRange最終行  現時点ではこれを対応
    # UsedRange 指定行-UsedRange最終行
    # ...
    @sheet.UsedRange.Columns(column_name_or_number).Value.flatten
  end

  def autofiltermode
    @sheet.AutoFilterMode
  end

  def autofilter(sw)
    # line 1 autofilter

    if autofiltermode
      @sheet.Rows(1).AutoFilter # Once Un AutoFilter
      @sheet.Rows(1).AutoFilter if sw # Re AutoFilter
    else
      @sheet.Rows(1).AutoFilter if sw
    end
  end

  def showalldata
    # http://club-vba.tokyo/vba-showalldata/
    @sheet.ShowAllData if @sheet.FilterMode
  end

  def columns_autofit(column = nil)
    # https://excelwork.info/excel/cellautofit/
    column ? @sheet.Columns(column).AutoFit : @sheet.Columns.AutoFit
  end

  def range2picture(range,pict_width = nil,pict_height = nil,pos_top = nil, pos_left = nil)
    # make cell range to picture
    @sheet.Range(range).CopyPicture
    @sheet.Pictures.Paste.Select

    # resize
    @sheet.Shapes(1).Width  = pict_width  if pict_width
    @sheet.Shapes(1).Height = pict_height if pict_height

    # move picture to specified position
    @sheet.Shapes(1).Top  = pos_top  if pos_top
    @sheet.Shapes(1).Left = pos_left if pos_left
  end

end
