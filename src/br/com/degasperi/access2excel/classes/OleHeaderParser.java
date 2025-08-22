package br.com.degasperi.access2excel.classes;

import java.util.Arrays;

/**
 * A parser for extracting metadata from MS Access OLE object headers.
 * This class is designed to be robust against malformed OLE data to prevent crashes.
 */
public class OleHeaderParser {
	private static int OBJECT_SIZE_INDEX = 8;
	private static int CLASS_SIZE_INDEX = 10;
	private static int OBJECT_POS_INDEX = 12;
	private static int CLASS_POS_INDEX = 14;
	private String objectName;
	private String className;

	/**
	 * Gets the object name extracted from the OLE header.
	 * @return The object name, or an error string if parsing failed.
	 */
	public String getObjectName() {
		return objectName;
	}

	/**
	 * Gets the class name extracted from the OLE header.
	 * @return The class name, or an error string if parsing failed.
	 */
	public String getClassName() {
		return className;
	}

	private static int getInt(byte[] byteArray, int position) {
		return byteArray[position + 1] << 8 & 0xFF00 | byteArray[position] & 0xFF;
	}

	/**
	 * Parses the byte array of an OLE object to extract its name and class.
	 * This constructor is safe to use with untrusted data and will not throw exceptions
	 * for malformed headers.
	 *
	 * @param oleBytes The byte array representing the OLE object.
	 */
	public OleHeaderParser(byte[] oleBytes) {
		try {
			if (oleBytes == null || oleBytes.length < 16) { // Header must be at least 16 bytes
				objectName = "Invalid OLE Header";
				className = "Invalid OLE Header";
				return;
			}

			int objectSize = getInt(oleBytes, OBJECT_SIZE_INDEX);
			int classSize = getInt(oleBytes, CLASS_SIZE_INDEX);
			int objectPos = getInt(oleBytes, OBJECT_POS_INDEX);
			int classPos = getInt(oleBytes, CLASS_POS_INDEX);

			if (objectPos < 0 || objectSize <= 0 || objectPos + objectSize > oleBytes.length || classPos < 0
					|| classSize <= 0 || classPos + classSize > oleBytes.length) {
				objectName = "Invalid OLE Data";
				className = "Invalid OLE Data";
				return;
			}

			objectName = new String(Arrays.copyOfRange(oleBytes, objectPos, objectPos + objectSize - 1));
			className = new String(Arrays.copyOfRange(oleBytes, classPos, classPos + classSize - 1));
		} catch (Exception e) {
			// Catch any other unexpected errors during parsing
			objectName = "OLE Parse Error";
			className = "OLE Parse Error";
		}
	}
}
