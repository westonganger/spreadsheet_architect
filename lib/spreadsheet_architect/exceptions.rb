module SpreadsheetArchitect
  module Exceptions
    class NoDataError < StandardError
      def initialize
        super("Missing or incorrect :data option")
      end
    end

    class NoInstancesError < StandardError
      def initialize
        super("Missing :instances option")
      end
    end
    
    class SpreadsheetColumnsNotDefined < StandardError
      def initialize(klass=nil)
        super("The spreadsheet_columns option is not defined on #{klass.name}")
      end
    end
  end
end
