package br.com.degasperi.access2excel.classes;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;

public class Arguments {
	private String inputFile;
	private String outputFile;
	private String outputFormat = "XLSX";
	private static final String RX_PARAM = "^(\\-|\\/)(?<name>[^\\=]+)(\\=(?<value>[^\\ ]+))?$";

	public Arguments(String[] args) throws Exception {
		try {
			for (String arg : args) {
				Pattern pattern = Pattern.compile(RX_PARAM);
				Matcher matcher = pattern.matcher(arg);
				if (!matcher.find())
					throw new ArgumentException(ArgumentExceptionType.INVALID, arg);

				String name = matcher.group("name");
				String value = matcher.group("value");
				if ("inputfile".equalsIgnoreCase(name))
					inputFile = value;
				else if ("outputfile".equalsIgnoreCase(name))
					outputFile = value;
				else if ("format".equalsIgnoreCase(name)) {
					if (!"xls".equalsIgnoreCase(value) && !"xlsx".equalsIgnoreCase(value))
						throw new ArgumentException(ArgumentExceptionType.INVALID, arg);
					outputFormat = value.toUpperCase();
				} else
					throw new ArgumentException(ArgumentExceptionType.UNEXPECTED, arg);
			}

			if (StringUtils.isEmpty(inputFile))
				throw new ArgumentException(ArgumentExceptionType.EXPECTED, "inputFile");
		} catch (Exception e) {
			StringBuilder sb = new StringBuilder();
			sb.append(e.getMessage());
			sb.append("\n");
			sb.append(
					"Use:\n\tjava -jar Access2Excel.jar -inputFile=<inputFile> [-outputFile=<outputFile>] [-format=<format>]\n");
			sb.append("where\n\t<inputFile>: A file name to Access database (.MDB or .ACCDB)");
			sb.append(
					"\n\t<outputFile>: A file name to Excel streadsheet (optional, default \"inputFile.xls\" or \"inputFile.xlsx\")");
			sb.append("\n\t<format>: Output format (\"XLS\" or \"XLSX\", optional, default \"XLSX\")");
			System.err.println(sb.toString());
			System.exit(0);
		}
	}

	public String getInputFile() {
		return inputFile;
	}

	public String getOutputFile() {
		if (StringUtils.isEmpty(outputFile))
			return FilenameUtils.removeExtension(inputFile) + "." + outputFormat.toLowerCase();
		else
			return outputFile;
	}

	public String getOutputFormat() {
		return outputFormat;
	}

	private static enum ArgumentExceptionType {
		INVALID, UNEXPECTED, EXPECTED;
	}

	private static class ArgumentException extends Exception {
		private static final long serialVersionUID = -5117423072929460460L;

		public ArgumentException(ArgumentExceptionType type, String arg) {
			super(String.format("Argument \"%s\" %s", arg, type.toString().toLowerCase()));
		}

	}
}
