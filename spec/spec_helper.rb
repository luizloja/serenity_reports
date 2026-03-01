$:.unshift(File.join(File.dirname(__FILE__), "..", "lib"))

require 'serenity_report'
require 'fileutils'

Dir[File.dirname(__FILE__) + "/support/**/*.rb"].each { |f| require f }

Ship = Struct.new(:name, :type)
Person = Struct.new(:name, :skill, :col1, :col2, :col3)

TMP_DIR = File.expand_path('../tmp', __dir__)

def fixture(name)
  File.join(File.dirname(__FILE__), '..', 'fixtures', name)
end

def tmp(name)
  File.join(TMP_DIR, name)
end

RSpec.configure do |config|
  config.filter_run focus: true
  config.run_all_when_everything_filtered = true

  config.before(:suite) do
    FileUtils.mkdir_p(TMP_DIR)
  end

end
