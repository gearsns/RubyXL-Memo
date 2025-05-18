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

# テーブルを使う

## ソース

```Ruby
require 'rubyXL'
require 'rubyXL/convenience_methods'

module RubyXL
    # TableFileを出力するように
    class Worksheet
        define_relationship(RubyXL::TableFile, :target_file)
        def related_objects
          comments + printer_settings + [target_file]
        end
    end

    # Table関連宣言

    class TableColumn < OOXMLObject
        define_attribute(:id, :int)
        define_attribute(:name, :string)
        define_element_name 'tableColumn'
    end

    class TableColumns < OOXMLContainerObject
        define_child_node(RubyXL::TableColumn, :collection => :with_count)
        define_element_name 'tableColumns'
    end

    class TableStyleInfo < OOXMLObject
        define_attribute(:name, :string)
        define_attribute(:showColumnStripes, :bool)
        define_attribute(:showFirstColumn, :bool)
        define_attribute(:showLastColumn, :bool)
        define_attribute(:showRowStripes, :bool)
        define_element_name 'tableStyleInfo'
    end

    class Table < OOXMLTopLevelObject
        define_attribute(:id, :int)
        define_attribute(:name, :string)
        define_attribute(:displayName, :string)
        define_attribute(:ref, :string)
        define_attribute(:totalRowShown, :bool)
        define_child_node(RubyXL::AutoFilter)
        define_child_node(RubyXL::TableColumns)
        define_child_node(RubyXL::TableStyleInfo)
        define_element_name 'table'
        set_namespaces('http://schemas.openxmlformats.org/spreadsheetml/2006/main' => nil,
                       'http://schemas.openxmlformats.org/markup-compatibility/2006' => 'mc',
                       'http://schemas.microsoft.com/office/spreadsheetml/2014/revision' => 'xr',
                       'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3' => 'xr3',
                       )
    end
    class TablePart < OOXMLObject
        define_attribute(:'r:id', :string)
        define_element_name 'tablePart'
    end
end

def createTable(worksheet, header, data)
    # Add relationship
    worksheet.relationship_container ||= RubyXL::OOXMLRelationshipsFile.new
    relationships = worksheet.relationship_container.relationships
    tableNo = relationships.size + 1
    tableRefId = "rId#{tableNo}"
    relationships << RubyXL::Relationship.new(
        :id => tableRefId, 
        :target => "../tables/table#{tableNo}.xml",
        :type => RubyXL::TableFile::REL_TYPE
    )
    # Table
    rangeReference = "A1:#{ind2ref(data.size + 1 - 1, header.size - 1)}"
    table = RubyXL::Table.new(
        :id => tableNo, 
        :name => "Table#{tableNo}",
        :display_name => "Table#{tableNo}",
        :ref => rangeReference,
        :total_row_shown => false
    )
    table.table_style_info = RubyXL::TableStyleInfo.new(
        :name => "TableStyleMedium2",
        :show_first_column => false,
        :show_last_column => false,
        :show_row_stripes => true,
        :show_column_stripes => false
    )
    table.auto_filter = RubyXL::AutoFilter.new(:ref => rangeReference)
    table.table_columns = RubyXL::TableColumns.new
    # Set Data
    header.each_with_index{|n, i|
        worksheet.add_cell(0, i, n)
        worksheet.workbook.shared_strings_container.add(n)
        table.table_columns << RubyXL::TableColumn.new(:id => i+1, :name => header[i])
    }
    data.each_with_index{|line, row|
        line.each_with_index{|n, i|
            worksheet.add_cell(row+1, i, n)
        }
    }
    # Set table
    worksheet.target_file = RubyXL::TableFile.new(::Pathname.new('/').join('xl', 'tables', "table#{tableNo}.xml"), table.write_xml)
    worksheet.table_parts = RubyXL::TableParts.new
    worksheet.table_parts << RubyXL::RID.new(:r_id => tableRefId)
end

def ind2ref(row = 0, col = 0)
    str = ''

    loop do
        x = col % 26
        str = ('A'.ord + x).chr + str
        col = (col / 26).floor - 1
        break if col < 0
    end

    str += (row + 1).to_s
end

# テストデータ

workbook = RubyXL::Workbook.new
worksheet = workbook[0]
header = [
    "a", "b", "c"
]
data = [
    [1,2,3],
    [4,5,6],
    [7,8,9],
]

createTable(worksheet, header, data)

workbook.stream.flush
workbook.stream.close
workbook.write("./table.xlsx")

```