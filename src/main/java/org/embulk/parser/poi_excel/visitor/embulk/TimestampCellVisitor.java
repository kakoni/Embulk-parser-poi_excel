package org.embulk.parser.poi_excel.visitor.embulk;

import java.time.Instant;
import java.util.Date;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;
import org.embulk.spi.type.TimestampType;
import org.embulk.util.config.Config;
import org.embulk.util.config.ConfigDefault;
import org.embulk.util.config.Task;
import org.embulk.util.config.units.ColumnConfig;
import org.embulk.util.config.units.SchemaConfig;
import org.embulk.util.timestamp.TimestampFormatter;
import java.time.format.DateTimeParseException;
import java.util.Optional;

import static org.embulk.parser.poi_excel.PoiExcelParserPlugin.getConfigMapper;

public class TimestampCellVisitor extends CellVisitor {

	public TimestampCellVisitor(PoiExcelVisitorValue visitorValue) {
		super(visitorValue);
	}

	@Override
	public void visitCellValueNumeric(Column column, Object source, double value) {
		Date date = DateUtil.getJavaDate(value);
		Instant instant = date.toInstant();
		pageBuilder.setTimestamp(column, instant);
	}

	@Override
	public void visitCellValueString(Column column, Object source, String value) {
		Instant instant;
		try {
			TimestampFormatter formatter = getTimestampFormatter(column);
            instant = formatter.parse(value);
		} catch (DateTimeParseException e) {
			doConvertError(column, value, e);
			return;
		}

		pageBuilder.setTimestamp(column, instant);
	}

	@Override
	public void visitCellValueBoolean(Column column, Object source, boolean value) {
		doConvertError(column, value, new UnsupportedOperationException(
				"unsupported conversion Excel boolean to Embulk timestamp"));
	}

	@Override
	public void visitCellValueError(Column column, Object source, int code) {
		doConvertError(column, code, new UnsupportedOperationException(
				"unsupported conversion Excel Cell error code to Embulk timestamp"));
	}

	@Override
	public void visitValueLong(Column column, Object source, long value) {
		pageBuilder.setTimestamp(column, Instant.ofEpochMilli(value));
	}

	@Override
	public void visitSheetName(Column column) {
		Sheet sheet = visitorValue.getSheet();
		visitSheetName(column, sheet);
	}

	@Override
	public void visitSheetName(Column column, Sheet sheet) {
		doConvertError(column, sheet.getSheetName(), new UnsupportedOperationException(
				"unsupported conversion sheet_name to Embulk timestamp"));
	}

	@Override
	public void visitRowNumber(Column column, int index1) {
		doConvertError(column, index1, new UnsupportedOperationException(
				"unsupported conversion row_number to Embulk timestamp"));
	}

	@Override
	public void visitColumnNumber(Column column, int index1) {
		doConvertError(column, index1, new UnsupportedOperationException(
				"unsupported conversion column_number to Embulk timestamp"));
	}

	@Override
	protected void doConvertErrorConstant(Column column, String value) throws Exception {
		TimestampFormatter formatter = getTimestampFormatter(column);
		pageBuilder.setTimestamp(column, formatter.parse(value));
	}

	private TimestampFormatter[] timestampFormatters;

	protected final TimestampFormatter getTimestampFormatter(Column column) {
		if (timestampFormatters == null) {
			PluginTask task = visitorValue.getPluginTask();
			timestampFormatters = newTimestampColumnFormattersForParsing(task, task.getColumns());
		}
		return timestampFormatters[column.getIndex()];
	}

	public static TimestampFormatter[] newTimestampColumnFormattersForParsing(final PluginTask task, final SchemaConfig schema) {
		final TimestampFormatter[] formatters = new TimestampFormatter[schema.getColumnCount()];
		int i = 0;
		for (final ColumnConfig column : schema.getColumns()) {
			if (column.getType() instanceof TimestampType) {
				final TimestampColumnOptionForParsing columnOption = getConfigMapper().map(column.getOption(), TimestampColumnOptionForParsing.class);

				final String pattern = columnOption.getFormat().orElse(task.getDefaultTimestampFormat());
				formatters[i] = TimestampFormatter.builder(pattern, true)
						.setDefaultZoneFromString(columnOption.getTimeZoneId().orElse(task.getDefaultTimeZoneId()))
						.setDefaultDateFromString(columnOption.getDate().orElse(task.getDefaultDate()))
						.build();
			}
			i++;
		}
		return formatters;
	}

	private interface TimestampColumnOptionForParsing extends Task {
		@Config("timezone")
		@ConfigDefault("null")
		Optional<String> getTimeZoneId();

		@Config("format")
		@ConfigDefault("null")
		Optional<String> getFormat();

		@Config("date")
		@ConfigDefault("null")
		Optional<String> getDate();
	}
}
