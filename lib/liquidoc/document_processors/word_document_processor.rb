module Liquidoc
  module DocumentProcessors
    class WordDocumentProcessor < BaseDocumentProcessor
      CONTENT_ENTRY_NAME = 'word/document.xml'

      def_delegator :template, :errors
      def_delegator :template, :tags
      def_delegator :template, :variables
      attr_accessor :current_file
      
      def render(context = {})
        Zip::OutputStream.write_buffer(StringIO.new) do |out|
          document_entries.each do |document_entry|
            out.put_next_entry document_entry.name

            if document_entry.name.include?("header") or document_entry.name.include?("footer") or document_entry.name == 'word/document.xml'
              self.current_file = document_entry.name
              begin
                out.write template.render(context)
              rescue
                out.write document_entries.read(document_entry.name)
              end
            else
              out.write document_entries.read(document_entry.name)
            end
          end
        end.string
      end

      def content
        @content = document_entries.read(self.current_file)
      end

      def sanitizer
        @sanitizer = Sanitizers::WordSanitizer.new(content)
      end

      def template
        @template = Template.new(sanitizer.document)
      end

      def clean_content
        @clean_content ||= sanitizer.clean_content
      end

      private

      def document_entries
        return @document_entries if defined?(@document_entries)

        @document_entries = Zip::File.open(path)
        @document_entries.close
        @document_entries
      end
    end
  end
end
