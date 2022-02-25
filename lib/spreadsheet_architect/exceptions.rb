module SpreadsheetArchitect
  module Exceptions

    class OptionTypeError < TypeError
      def initialize(which)
        super("Invalid data type for the #{which} option")
      end
    end

    class ArgumentError < ::ArgumentError
      def initialize(msg)
        super(msg)
      end
    end

    class InvalidRangeError < ArgumentError
      def initialize(type, range, message: nil)
        default_msg = "Invalid range `#{range}` passed"

        if message.nil?
          case type
          when :missing_range_keys
            message = "Missing :rows or :columns key"
          when :format
            message = "Format must be as follows: A1:D4"
          when :type
            message = "Valid types are Hash and String"
          end
        end

        super([default_msg, message].compact.join(". "))
      end
    end

    class InvalidRangeValue < ArgumentError
      def initialize(type, range)
        default_msg = "Invalid range `#{range}` passed. Some of the :#{type} specified are either an invalid value or are greater than the total number of #{type}"
      end
    end

    class InvalidColumnError < ArgumentError
      def initialize(col)
        super("Invalid Column `#{col}` given for column_types options")
      end
    end

    class InvalidRangeOptionError < ArgumentError
      def initialize(key, opt)
        default_msg = "Invalid :#{key} option for `#{opt}`"

        case key
        when :columns, :rows
          message = ":#{key} can be an integer, range, or :all"
        end

        super([default_msg, message].compact.join(". "))
      end
    end

    class NoDataError < ArgumentError
      def initialize
        super("Missing :data or :instances option")
      end
    end

    class MultipleDataSourcesError < ArgumentError
      def initialize
        super("Both :data and :instances options cannot be combined, please choose one.")
      end
    end
    
    class SpreadsheetColumnsNotDefinedError < ArgumentError
      def initialize(klass, method_name='spreadsheet_columns')
        super("The instance method `#{method_name}` is not defined on #{klass}")
      end
    end
  end
end
