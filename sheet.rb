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
                 xaxistitle:  nil, yaxistitle:  nil)
    # height: 428, width: 907 to fit Power Point slide area height: 15.1cm,width 32.31cm
    # 17.64cm 500pt -> 1cm = 28.354pt

    range_obj.Select # needed as for chart SetSourceData

    shapes = @sheet.Shapes # 指定されたシートのすべての**Shape** オブジェクトのコレクションです。
    # AddChart2 ドキュメントにグラフを追加します。 グラフを表す**Shape** オブジェクトを返し、指定されたコレクションに追加します。
    graph_shape = shapes.AddChart2(Style: -1,XlChartType: charttype,Left: left,Top: top,Width: width,Height: height)

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

end
