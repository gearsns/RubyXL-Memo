# 条件付き書式を使う
RubyXL で 条件付き書式 を設定しようとして方法が探しても見つからなかったので

## ソース

```Ruby
require 'rubygems'
require 'rubyxl'

# 条件付き書式 サンプル
workbook = RubyXL::Workbook.new
worksheet = workbook[0]

unless workbook.stylesheet
	workbook.stylesheet = RubyXL::Stylesheet.new
	# https://github.com/weshatheleopard/rubyXL/blob/0a42ab5bad5e9be796cc0f7710d58af9aabcb7cc/lib/rubyXL/objects/stylesheet.rb#L158
end
unless workbook.stylesheet.dxfs
	workbook.stylesheet.dxfs = RubyXL::DXFs.new
	# https://github.com/weshatheleopard/rubyXL/blob/0a42ab5bad5e9be796cc0f7710d58af9aabcb7cc/lib/rubyXL/objects/stylesheet.rb#L111
end

# 使用するスタイルを登録
# ※文字を赤色にする
# 他の書式を設定する際には、DXFの定義を参照
# https://github.com/weshatheleopard/rubyXL/blob/0a42ab5bad5e9be796cc0f7710d58af9aabcb7cc/lib/rubyXL/objects/stylesheet.rb#L99
workbook.stylesheet.dxfs << RubyXL::DXF.new(:font => RubyXL::Font.new(:color => RubyXL::Color.new(:rgb => '00FF00FF')))

# ルールを登録
# ※セルの値が、0よりも小さい時に上記で登録したスタイルを適用
worksheet.conditional_formatting = RubyXL::ConditionalFormatting.new(:sqref => "A1:A2", 
:cf_rule => RubyXL::ConditionalFormattingRule.new(
	   :type => "expression",
	   :priority => "2",
	   :dxf_id => workbook.stylesheet.dxfs.count - 1,
	   :formulas => [
		   RubyXL::Formula.new(
			   :t => 'formula',
			   :expression => '=IF(A1<0,TRUE,FALSE)'
		   )
	   ]
   )
)

# テストデータ
worksheet.add_cell(0, 0, -1)
worksheet.add_cell(1, 0, 1)
workbook.write 'conditionalFormattingRule.xlsx'
```
