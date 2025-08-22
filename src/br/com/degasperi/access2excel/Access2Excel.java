package br.com.degasperi.access2excel;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

import br.com.degasperi.access2excel.classes.Access2Excel.OutputFormat;

public class Access2Excel {

	/**
	 * Main entry point for the Access2Excel command-line tool.
	 * <p>
	 * Parses command-line arguments and triggers the database conversion using Apache Commons CLI.
	 *
	 * @param args The command-line arguments.
	 */
	public static void main(String[] args) {
		Options options = new Options();

		Option input = new Option("i", "inputFile", true, "input file path");
		input.setRequired(true);
		options.addOption(input);

		Option output = new Option("o", "outputFile", true, "output file path");
		output.setRequired(false);
		options.addOption(output);

		Option format = new Option("f", "format", true, "output format (XLS or XLSX)");
		format.setRequired(false);
		options.addOption(format);

		options.addOption(new Option("h", "help", false, "print this message"));

		CommandLineParser parser = new DefaultParser();
		HelpFormatter formatter = new HelpFormatter();
		CommandLine cmd;

		try {
			cmd = parser.parse(options, args);
		} catch (ParseException e) {
			System.err.println(e.getMessage());
			formatter.printHelp("Access2Excel", options);
			return;
		}

		if (cmd.hasOption("help")) {
			formatter.printHelp("Access2Excel", options);
			return;
		}

		try (br.com.degasperi.access2excel.classes.Access2Excel access2Excel = new br.com.degasperi.access2excel.classes.Access2Excel()) {
			String inputFile = cmd.getOptionValue("inputFile");
			String outputFile = cmd.getOptionValue("outputFile");
			String outputFormat = cmd.getOptionValue("format");

			if (outputFormat != null) {
				try {
					access2Excel.setOutputFormat(OutputFormat.valueOf(outputFormat.toUpperCase()));
				} catch (IllegalArgumentException e) {
					System.err.println("Invalid format: " + outputFormat + ". Use XLS or XLSX.");
					return;
				}
			}

			access2Excel.open(inputFile);
			access2Excel.convert(outputFile);

		} catch (Exception e) {
			System.err.println("Error: " + e.getClass().getSimpleName() + (e.getMessage() == null ? "" : " - " + e.getMessage()));
		}
	}

}
