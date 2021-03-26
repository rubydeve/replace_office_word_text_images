require 'open-uri'
require 'zip'

require './to_html/caracal/core/bookmarks'
require './to_html/caracal/core/custom_properties'
require './to_html/caracal/core/file_name'
require './to_html/caracal/core/fonts'
require './to_html/caracal/core/iframes'
require './to_html/caracal/core/ignorables'
require './to_html/caracal/core/images'
require './to_html/caracal/core/list_styles'
require './to_html/caracal/core/lists'
require './to_html/caracal/core/namespaces'
require './to_html/caracal/core/page_breaks'
require './to_html/caracal/core/page_numbers'
require './to_html/caracal/core/page_settings'
require './to_html/caracal/core/relationships'
require './to_html/caracal/core/rules'
require './to_html/caracal/core/styles'
require './to_html/caracal/core/tables'
require './to_html/caracal/core/text'

require './to_html/caracal/renderers/app_renderer'
require './to_html/caracal/renderers/content_types_renderer'
require './to_html/caracal/renderers/core_renderer'
require './to_html/caracal/renderers/custom_renderer'
require './to_html/caracal/renderers/document_renderer'
require './to_html/caracal/renderers/fonts_renderer'
require './to_html/caracal/renderers/footer_renderer'
require './to_html/caracal/renderers/numbering_renderer'
require './to_html/caracal/renderers/package_relationships_renderer'
require './to_html/caracal/renderers/relationships_renderer'
require './to_html/caracal/renderers/settings_renderer'
require './to_html/caracal/renderers/styles_renderer'


module Caracal
  class Document

    #------------------------------------------------------
    # Configuration
    #------------------------------------------------------

    # mixins (order is important)
    include Caracal::Core::CustomProperties
    include Caracal::Core::FileName
    include Caracal::Core::Ignorables
    include Caracal::Core::Namespaces
    include Caracal::Core::Relationships

    include Caracal::Core::Fonts
    include Caracal::Core::PageSettings
    include Caracal::Core::PageNumbers
    include Caracal::Core::Styles
    include Caracal::Core::ListStyles

    include Caracal::Core::Bookmarks
    include Caracal::Core::IFrames
    include Caracal::Core::Images
    include Caracal::Core::Lists
    include Caracal::Core::PageBreaks
    include Caracal::Core::Rules
    include Caracal::Core::Tables
    include Caracal::Core::Text


    #------------------------------------------------------
    # Public Class Methods
    #------------------------------------------------------

    #============ OUTPUT ==================================

    # This method renders a new Word document and returns it as a
    # a string.
    #
    def self.render(f_name = nil, &block)
      docx   = new(f_name, &block)
      buffer = docx.render

      buffer.rewind
      buffer.sysread
    end

    # This method renders a new Word document and saves it to the
    # file system.
    #
    def self.save(f_name = nil, &block)
      docx   = new(f_name, &block)
      docx.save
      # buffer = docx.render
      #
      # File.open(docx.path, 'wb') { |f| f.write(buffer.string) }
    end



    #------------------------------------------------------
    # Public Instance Methods
    #------------------------------------------------------

    # This method instantiates a new word document.
    #
    def initialize(name = nil, &block)
      file_name(name)

      page_size
      page_margins top: 1440, bottom: 1440, left: 1440, right: 1440
      page_numbers

      [:font, :list_style, :namespace, :relationship, :style].each do |method|
        collection = self.class.send("default_#{ method }s")
        collection.each do |item|
          send(method, item)
        end
      end

      if block_given?
        (block.arity < 1) ? instance_eval(&block) : block[self]
      end
    end


    #============ GETTERS =================================

    # This method returns an array of models which constitute the
    # set of instructions for producing the document content.
    #
    def contents
      @contents ||= []
    end


    #============ RENDERING ===============================

    # This method renders the word document instance into
    # a string buffer. Order is important!
    #
    def render
      buffer = ::Zip::OutputStream.write_buffer do |zip|
        render_package_relationships(zip)
        render_content_types(zip)
        render_app(zip)
        render_core(zip)
        render_custom(zip)
        render_fonts(zip)
        render_footer(zip)
        render_settings(zip)
        render_styles(zip)
        render_document(zip)
        render_relationships(zip)   # Must go here: Depends on document renderer
        render_media(zip)           # Must go here: Depends on document renderer
        render_numbering(zip)       # Must go here: Depends on document renderer
      end
    end


    #============ SAVING ==================================

    def save
      buffer = render

      File.open(path, 'wb') { |f| f.write(buffer.string) }
    end


    #------------------------------------------------------
    # Private Instance Methods
    #------------------------------------------------------
    private

    #============ RENDERERS ===============================

    def render_app(zip)
      content = ::Caracal::Renderers::AppRenderer.render(self)

      zip.put_next_entry('docProps/app.xml')
      zip.write(content)
    end

    def render_content_types(zip)
      content = ::Caracal::Renderers::ContentTypesRenderer.render(self)

      zip.put_next_entry('[Content_Types].xml')
      zip.write(content)
    end

    def render_core(zip)
      content = ::Caracal::Renderers::CoreRenderer.render(self)

      zip.put_next_entry('docProps/core.xml')
      zip.write(content)
    end

    def render_custom(zip)
      content = ::Caracal::Renderers::CustomRenderer.render(self)

      zip.put_next_entry('docProps/custom.xml')
      zip.write(content)
    end

    def render_document(zip)
      content = ::Caracal::Renderers::DocumentRenderer.render(self)

      zip.put_next_entry('word/document.xml')
      zip.write(content)
    end

    def render_fonts(zip)
      content = ::Caracal::Renderers::FontsRenderer.render(self)

      zip.put_next_entry('word/fontTable.xml')
      zip.write(content)
    end

    def render_footer(zip)
      content = ::Caracal::Renderers::FooterRenderer.render(self)

      zip.put_next_entry('word/footer1.xml')
      zip.write(content)
    end

    def render_media(zip)
      images = relationships.select { |r| r.relationship_type == :image }
      images.each do |rel|
        if rel.relationship_data.to_s.size > 0
          content = rel.relationship_data
        else
          content = open(rel.relationship_target).read
        end

        zip.put_next_entry("word/#{ rel.formatted_target }")
        zip.write(content)
      end
    end

    def render_numbering(zip)
      content = ::Caracal::Renderers::NumberingRenderer.render(self)

      zip.put_next_entry('word/numbering.xml')
      zip.write(content)
    end

    def render_package_relationships(zip)
      content = ::Caracal::Renderers::PackageRelationshipsRenderer.render(self)

      zip.put_next_entry('_rels/.rels')
      zip.write(content)
    end

    def render_relationships(zip)
      content = ::Caracal::Renderers::RelationshipsRenderer.render(self)

      zip.put_next_entry('word/_rels/document.xml.rels')
      zip.write(content)
    end

    def render_settings(zip)
      content = ::Caracal::Renderers::SettingsRenderer.render(self)

      zip.put_next_entry('word/settings.xml')
      zip.write(content)
    end

    def render_styles(zip)
      content = ::Caracal::Renderers::StylesRenderer.render(self)

      zip.put_next_entry('word/styles.xml')
      zip.write(content)
    end

  end
end
