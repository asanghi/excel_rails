require 'spreadsheet'

module Spreadsheet
  module Rails
    module SpreadsheetHelper

      def excel_document(opts={})
        download = opts.delete(:force_download)
        filename = opts.delete(:filename)
        workbook = Spreadsheet::Workbook.new
        yield(workbook)
        disposition(download, filename) if (download || filename)
        workbook
      end

      def disposition(download, filename)
        download = true if (filename && download == nil)
        disposition = download ? "attachment;" : "inline;"
        disposition += " filename=#{filename}" if filename
        headers["Content-Disposition"]=disposition
      end
    end

    class TemplateHandler < ActionView::Template::Handler
      self.default_format = :xls
      def self.call(template)
        "sio = StringIO.new; #{template.source.strip}.write(sio); sio.string"
      end
    end
  end
end

Mime::Type.register_alias "application/xls", :xls
ActionView::Template.register_template_handler(:rxls, Spreadsheet::Rails::TemplateHandler)
ActionView::Base.send(:include, Spreadsheet::Rails::SpreadsheetHelper)
