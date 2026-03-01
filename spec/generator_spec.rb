require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

module SerenityReport
  describe Generator do
    it 'makes context from instance variables and runs the provided template' do
      class GeneratorClient
        include SerenityReport::Generator

        def generate_odt
          @crew = ['Mal', 'Inara', 'Wash', 'Zoe']

          render_odt fixture('odt/loop.odt')
        end
      end

      client = GeneratorClient.new
      expect { client.generate_odt }.not_to raise_error
      expect(tmp('loop_output.odt')).to be_a_document
    end
  end
end
