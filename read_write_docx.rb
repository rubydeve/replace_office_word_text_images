class Customer
  require 'docx'
  require 'base64'
  require './to_html/caracal.rb'
  require 'open-uri'
  def initialize(file_path, hash)
    @file_path = file_path
    @hash = hash
    @map_data = []
    @download_images = []
  end

  def read
    doc = Docx::Document.open(@file_path)
    doc.paragraphs.each do |p|
      @map_data.push({name: p.text}) if !p.text.strip.empty?
    end
    doc.tables.each do |table|
      table.rows.each do |row| # Row-based iteration
        row.cells.each do |cell|
          @map_data.push({name: cell.text,type: "table"}) if !cell.text.strip.empty?
        end
      end
    end
  end

  def write
    Caracal::Document.save 'example.docx' do |docx|
      doc = Docx::Document.open(@file_path)
      doc.paragraphs.each do |p|
        @hash.each do |i| 
          # if text match with {example} == {example} 
          if !p.text.strip.empty? and "{#{i[:name]}}" == p.text
            font_name = set_font_style(p.node)
            create_tag(i , docx,p.font_size, font_name)
          end
        end
      end
      doc.tables.each do |table|
        table.rows.each do |row| 
          row.cells.each do |cell|
            @hash.each do |i| 
              if !cell.text.strip.empty? and "{#{i[:name]}}".downcase == cell.text.downcase and i[:type] =="table"
                create_tag(i ,docx,doc.document_properties[:font_size])
              end
            end
          end
        end
      end
    end
    clear_images
  end
  def set_font_style(tag_node)
    return tag_node.xpath("//w:rFonts").first.attributes.values
  end
  def create_tag(hash_index, docx,size ,font_name= nil)
    case hash_index[:type]
    when "text"
      docx.p hash_index[:content] ,size: (size * 2),font: font_name
    when "image"
      file_name = download_image(hash_index[:content])
      docx.img "./#{file_name}", width: 250, height: 100
    when "table"
      relpace_with_cell_image(hash_index[:content])
      docx.table hash_index[:content] do 
        index = 0
        hash_index[:content].each do |i|
          cell_style rows[index],size: (size * 2)  
          index += 1 
        end
      end
    end
  end

  def maper
    @map_data
  end

  def download_image(path)
    download = URI.open(path)
    file_name = download.base_uri.to_s.split('/')[-1]
    IO.copy_stream(download, "./#{file_name}")
    @download_images.push(file_name)
    return file_name
  end

  def relpace_with_cell_image(content)
    index = 0
    content.each do |tr|
      index2 = 0
      tr.each do |td|
        if td.class == Hash
          file_name = download_image(td[:content])
          cell  = Caracal::Core::Models::TableCellModel.new do
            img("./#{file_name}",width: 100, height: 100, align: 'center')
          end
          tr[index2] =  cell
        else
        end
        index2 += 1 
      end
      index += 1
    end
  end
  
  def clear_images
    @download_images.each do |file_name|
      File.delete("./#{file_name}")  if File.exist?("./#{file_name}")
    end
  end

  
  @customer = Customer.new("Doc1_before_script.docx",
    [
      {
        name: "nom",
        type: "text",
        content: "LEPIEZ"
      },
      {
        name: "photo",
        type: "image",
        content: "http://personal.psu.edu/xqz5228/jpg.jpg"
      },
      {
        name: "info",
        type: "table",
        style: "",
        content: [["firstname","Jérome"],["lastname","Bolomier"],["date de
        naissance","24/04/1977"],["",{content: "http://personal.psu.edu/xqz5228/jpg.jpg"}]]
      }
    ]
  )
  @customer.read
  @customer.maper
  @customer.write
end