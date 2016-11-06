package br.com.degasperi.access2excel.classes;

import java.util.Arrays;

public class OleHeaderParser {
	private static int OBJECT_SIZE_INDEX = 8;
	private static int CLASS_SIZE_INDEX = 10;
	private static int OBJECT_POS_INDEX = 12;
	private static int CLASS_POS_INDEX = 14;
	private String objectName;
	private String className;

	public String getObjectName() {
		return objectName;
	}

	public String getClassName() {
		return className;
	}

	private static int getInt(byte[] byteArray, int position) {
		return byteArray[position + 1] << 8 & 0xFF00 | byteArray[position] & 0xFF;
	}

	public OleHeaderParser(byte[] oleBytes) {
		int objectSize = getInt(oleBytes, OBJECT_SIZE_INDEX);
		int classSize = getInt(oleBytes, CLASS_SIZE_INDEX);
		int objectPos = getInt(oleBytes, OBJECT_POS_INDEX);
		int classPos = getInt(oleBytes, CLASS_POS_INDEX);
		objectName = new String(Arrays.copyOfRange(oleBytes, objectPos, objectPos + objectSize - 1));
		className = new String(Arrays.copyOfRange(oleBytes, classPos, classPos + classSize - 1));
	}
}
