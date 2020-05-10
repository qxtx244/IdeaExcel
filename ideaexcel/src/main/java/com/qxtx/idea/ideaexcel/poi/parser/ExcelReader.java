package com.qxtx.idea.ideaexcel.poi.parser;

import android.support.annotation.NonNull;
import android.util.Log;

import com.qxtx.idea.ideaexcel.poi.callback.IReadCallback;

import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.CellRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * Created in 2020/4/26 15:12
 *
 * @author QXTX-WORK
 * <p>
 * Description 支持读取和导出excel表格，支持xls, xlsx，csv三种格式的文件。
 *
 * <pre>
 * 目前的数据格式支持情况：
 * csv：
 * 0、由于是直接读取字符串，读取速度最快
 * 1、不支持单元格中使用换行；
 * 2、仅支持普通字符串；
 * 3、仅支持使用GBK编码格式的csv文件（windows下默认字符编码格式），
 *   使用其他字符编码（如使用Mac、Linux生成文件的默认字符编码为utf-8），可能会解析得到乱码（主要是中文字符）。
 *
 * xls：
 * 0、直接读取二进制数据，读取速度较快
 * 1、仅支持普通字符串，不支持日期（由于日期会被poi转化成double值，因此会和普通小数值混在一起无法区分）；
 *
 * xlsx：
 * 0、由于需要经历解压zip，读取xml数据等耗时操作，读取速度最慢
 * 1、仅支持普通字符串，不支持日期（由于日期会被poi转化成double值，因此会和普通小数值混在一起无法区分）；
 * 2、简单支持导出xlsx表格
 *
 * 注意：
 * 1、对于表格的读取，起始行/列序号为0
 *
 * </pre>
 */
public class ExcelReader {
    private volatile static ExcelReader instance;

    private final StringBuilder sb = new StringBuilder();

    /** 用于解析xls部分的内容读取回调监听器 */
    private FormatTrackingHSSFListener formatListener;

    private final String CSV_SEP = ",";
    private final char CSV_SEP_CHAR = ',';

    /**
     * <pre>
     * 用于处理单元格中包含csv逗号分隔符的情况
     * 单元格内有特殊字符的处理：用双引号将单元格内容包含起来，取值时需要去除双引号，有两种特殊符号：
     * ①仅处理带[,]的单元格：直接用""将整个单元格内容包含起来
     * ②仅处理带["]的单元格：用""将整个单元格内容包含起来，并且这个作为单元格数据的["]用两个连续的双引号表示。
     * ③处理[,]和["]都存在的单元格：同时使用①和②处理
     * 示例：
     * 单元格内容1：[,]      csv表示：[","]
     * 单元格内容2：[abc]    csv表示：[abc]
     * 单元格内容3：["]      csv表示：[""""]
     * 单元格内容4：[,abc"]  csv表示：[",abc"""]
     * 单元格内容5：[]  csv表示：[]
     * </pre>
     */
    private final char CSV_SPEC_CHAR = '"';

    /**
     * <pre>
     * 各种解析方案和对应的使用状态。
     * 键：后缀名称，对应一种解析方案
     * 值：解析方案的使用状态，见{@link SchemeState}。当某个方案的使用状态为{@link SchemeState#INVALID}，则表示此方案不可再次使用
     * </pre>
     */
    private final HashMap<String, Byte> mParserSchemeMap;

    /** 解析方案的状态 */
    @Retention(RetentionPolicy.SOURCE)
    public @interface SchemeState {
        /** 方案未被使用过，可尝试使用 */
        byte VALID = 0;
        /** 方案已经被使用过，不可重复使用 */
        byte INVALID = 0x1;
    }

    /** 文件名格式后缀 */
    @Retention(RetentionPolicy.SOURCE)
    public @interface Suffix {
        String XLS = ".xls";
        String XLSX = ".xlsx";
        String CSV = ".csv";
    }

    public static ExcelReader getInstance() {
        if (instance == null) {
            synchronized (ExcelReader.class) {
                if (instance == null) {
                    instance = new ExcelReader();
                }
            }
        }
        return instance;
    }

    private ExcelReader() {
        mParserSchemeMap = new HashMap<>();
        initSchemeState();
    }

    /**
     * 解析excel表格，支持多种格式
     *
     * 由于文件可能会被人为地重命名为不相符的格式后缀，因此通过直接识别后缀名去选择解析方案，可能会失败。
     *   在失败后继续尝试其他可用的解析方案。
     *
     * @param path 文件绝对路径
     * @param callback 给外部的事件回调
     */
    public void read(@NonNull String path, @NonNull IReadCallback callback) {
        File file = new File(path);
        if (!file.exists() || file.isDirectory()) {
            log("I", "非法文件");
            return ;
        }

        try {
            //直接通过后缀判断，容错率较低
            String suffix = path.substring(path.lastIndexOf(".")).toLowerCase();
            parseWithSuffix(suffix, file, callback);
        } catch (Exception e) {
            log("E", "解析表格发生异常：" + e.getLocalizedMessage());
            e.printStackTrace();
        }
    }

    /** 解析xls表格内容 */
    private boolean parseXls(@NonNull File file, @NonNull IReadCallback callback) throws Exception {
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));

        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(new XlsListener(callback));
        formatListener = new FormatTrackingHSSFListener(listener);

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        request.addListenerForAllRecords(formatListener);

        factory.processWorkbookEvents(request, fs);

        return true;
    }

    /**
     * 解析xlsx表格内容。
     * 如果无法以zip的形式读取文件，说明此文件不是xlsx，可能是直接将xls改名或者直接将csv改名，尝试使用这两种格式读取
     *
     * @return [true]已正确读取表格  [false]读取表格发生异常，可能需要尝试其他解析方案
     */
    private boolean parseXlsx(@NonNull File file, @NonNull IReadCallback callback) throws Exception {
        OPCPackage opcp = OPCPackage.open(file);
        ReadOnlySharedStringsTable table = new ReadOnlySharedStringsTable(opcp);
        XSSFReader reader = new XSSFReader(opcp);
        //读取表格内容
        return readSheet(reader, table, callback);
    }

    /**
     * 解析csv文件内容
     * @return [true]成功解析或解析失败但不需要更换其他解析方案  [false]解析失败，可能需要更换其他解析方案
     */
    private boolean parseCsv(@NonNull File file, @NonNull IReadCallback callback) throws Exception {
        if (!file.exists() || file.isDirectory()) {
            log("I", "文件不存在或者路径是一个目录");
            return true;
        }

        FileInputStream fis = new FileInputStream(file);
        InputStreamReader isr = new InputStreamReader(fis, Charset.forName("gbk"));
        //设置缓冲区大小为200KB，减少文件I/O次数
        BufferedReader br = new BufferedReader(isr, 200 * 1024);

        final String errMsg = "缺少\"，无法解析的单元格";
        int rowIndex = 0;
        String line;
        //遍历表格中的所有行
        while ((line = br.readLine()) != null) {
            List<String> rowInfo = new ArrayList<>();
            int startPos = 0;
            int lineLen = line.length();

            //遍历行数据，有三种情况：
            // 1、起始字符为["]，即意味着单元格内容包含了csv的特殊处理字符[,]或["]；
            //2、起始字符为[,]，即意味着已经处于单元格的结束位置（空单元格）；
            //3、起始字符为普通字符，即找到[,]即为单元格结束位置
            while (startPos < lineLen) {
                char startChar = line.charAt(startPos);
                //排除起始字符就是[,]的情况（说明跳过了一个空值）
                if (startChar == CSV_SEP_CHAR) {
                    log("I", "发现空值");
                    rowInfo.add("");
                    startPos++;
                    if (startPos == lineLen) {
                        //已经到达行尾了，当前行解析结束
                        rowInfo.add("");
                        break;
                    } else {
                        continue;
                    }
                }

                sb.delete(0, sb.length());

                String cell;
                //第一次碰到["]，为包裹单元格的左双引号
                if (startChar == CSV_SPEC_CHAR) {
                    sb.append(startChar);

                    //在此之后应至少还有两个字符
                    if (startPos + 2 > lineLen) {
                        throw new IllegalStateException(errMsg);
                    }

                    //在检索的过程中，将[""]还原为["]
                    for (startPos++; startPos < lineLen; startPos++) {
                        char c = line.charAt(startPos);

                        //第二次碰到["]
                        //下一个字符只能是["]或[,]，或者到达行结尾
                        //1、[,]：①[",]即单元格结束
                        //2、["]： ①如果已经到达行结尾，说明它属于右双引号且行解析已经结束；②如果未到行结尾，说明它属于单元格内容
                        if (c == CSV_SPEC_CHAR) {
                            int nextPos = startPos + 1;
                            if (nextPos == lineLen) {
                                //已经到达行结尾，结束
                                sb.append(c);
                                startPos++;
                                break;
                            }

                            char nextChar = line.charAt(nextPos);
                            if (nextChar == CSV_SEP_CHAR) {
                                sb.append(c);
                                startPos += 2;
                                break;
                            } else if (nextChar == CSV_SPEC_CHAR) {
                                //无论是到达行结尾还是当前字符属于单元格内容，都是跳过1个字符
                               sb.append(c);
                                startPos++;
                               continue;
                            } else {
                                throw new IllegalStateException("检测到非法数据，解析异常");
                            }

                        } else {
                            //非[,]，不需要额外处理
                            sb.append(c);
                        }
                    }

                    //注意需要去掉包裹单元格内容的双引号
                    cell = sb.toString();
                    cell = cell.substring(1, cell.length() - 1);
//                    log("I", "得到一个单元格数据：[" + cell + "].");
                } else {
                    int sepIndex = line.indexOf(CSV_SEP_CHAR, startPos);
                    cell = sepIndex == -1 ? line.substring(startPos) : line.substring(startPos, sepIndex);
                    startPos = sepIndex == -1 ? lineLen : (sepIndex + 1);
                }

                rowInfo.add(cell);

                //到达行结尾，如果前面是一个[,]，说明后面存在 一个空值，需要发现这个空值
                if (startPos == lineLen && line.charAt(startPos - 1) == CSV_SEP_CHAR) {
                    rowInfo.add("");
                }
            }

            callback.onRowRead(rowIndex, rowInfo);

            rowIndex++;
        }

        return true;
    }

    /**
     * 通过文件后缀名选择解析方案。当前方案解析失败时，自动尝试其他可用的方案
     * 目标解析结果：表格中每行数据拼接成一个List
     */
    private void parseWithSuffix(@NonNull String suffix, @NonNull File file, IReadCallback callback) {
        log("I", "开始解析：" + suffix + ",file=" + file.getPath());
        boolean isFinished;
        try {
            switch (suffix) {
                case Suffix.XLS:
                    isFinished = parseXls(file, callback);
                    break;
                case Suffix.XLSX:
                    isFinished = parseXlsx(file, callback);
                    break;
                case Suffix.CSV:
                    isFinished = parseCsv(file, callback);
                    break;
                default:
                    isFinished = false;
                    break;
            }
        } catch (Exception e) {
            log("E", "读取表格发生异常：" + e);
            e.printStackTrace();
            isFinished = false;
        }

        //尝试使用其他解析方案
        if (!isFinished) {
            log("I", "使用其他解析方案...");
            //当前采用的方案解析失败了，将其置为不可用状态。
            mParserSchemeMap.put(suffix, SchemeState.INVALID);

            for (String scheme : mParserSchemeMap.keySet()) {
                Byte state = mParserSchemeMap.get(scheme);
                //忽略不可用的解析方案
                if (state == SchemeState.INVALID) {
                    continue;
                }

                parseWithSuffix(scheme, file, callback);
            }
        } else {
            resetSchemeState();
        }
    }

    private void initSchemeState() {
        mParserSchemeMap.put(Suffix.XLS, SchemeState.VALID);
        mParserSchemeMap.put(Suffix.XLSX, SchemeState.VALID);
        mParserSchemeMap.put(Suffix.CSV, SchemeState.VALID);
    }

    private void resetSchemeState() {
        mParserSchemeMap.put(Suffix.XLS, SchemeState.INVALID);
        mParserSchemeMap.put(Suffix.XLSX, SchemeState.INVALID);
        mParserSchemeMap.put(Suffix.CSV, SchemeState.INVALID);
    }

    /**
     * 核心的读取表格数据方法
     * @return [true]已正确读取表格  [false]读取表格发生异常
     */
    private boolean readSheet(@NonNull XSSFReader reader,
                          @NonNull ReadOnlySharedStringsTable table,
                          @NonNull IReadCallback callback) {
        try {
            Iterator<InputStream> iterator = reader.getSheetsData();
            if (!iterator.hasNext()) {
                log("I", "可以正确服务表格，但未找到任何有效的sheet");
                //没有找到任何sheet，此时不用尝试其他解析方案了
                return true;
            }

            //只取第0张表格
            try (InputStream inputStream = iterator.next()) {
                XMLReader xmlReader = SAXHelper.newXMLReader();
                xmlReader.setContentHandler(new XSSFSheetXMLHandler(reader.getStylesTable(),
                        table, new XlsxSheetHandler(callback), false));
                xmlReader.parse(new InputSource(inputStream));
            }
        } catch (Exception e) {
            log("E", "读取表格内容发生异常：" + e.getLocalizedMessage());
            e.printStackTrace();
            return false;
        }

        return true;
    }

    private void log(String type, @NonNull String msg) {
        final String tag = "ExcelParser";
        if (type.toUpperCase().equals("E")) {
            Log.e(tag, msg);
        } else if (type.toUpperCase().equals("I")) {
            Log.i(tag, msg);
        } else {
            Log.d(tag, msg);
        }
    }

    /**
     * xssf读取表格文件，读取内容过程通过此对象回调出来
     * 目标解析结果：表格中每行数据拼接成一个List
     *
     * @see #parseXlsx(File, IReadCallback)
     */
    private final class XlsxSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        /** 表格行序号，从0开始计数 */
        private int curRow = 0;

        /** 表格每一行的列序号，每一行都从0开始计数列 */
        private int curColumn = 0;

        /** 给外部的回调 */
        private final IReadCallback callback;

        /** 一行数据 */
        private List<String> rowInfo;

        private XlsxSheetHandler(@NonNull IReadCallback callback) {
            this.callback = callback;
        }

        /**
         * 每一行数据读取开始，回调此方法
         */
        @Override
        public void startRow(int rowNum) {
            curRow = rowNum;
            rowInfo = new ArrayList<>();
        }

        /**
         * 每一行数据读取结束，回调此方法
         */
        @Override
        public void endRow(int rowNum) {
            curColumn = 0;

            callback.onRowRead(rowNum, rowInfo);
        }

        /**
         * 全部数据在此被回调
         * @param cellReference Excel中列的索引，列名+行名的字符串 如A1, B3，C2，A2……
         * @param formattedValue 单元格的数据，以字符串形式表示
         * @param comment 单元格的描述，现在还不需要关心
         */
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            rowInfo.add(formattedValue);

            curColumn++;
        }

        @Override
        public void headerFooter(String s, boolean b, String s1) {
//            log("I", "headerfooter!: s=" + s + ", b=" + b + ", s1=" + s1);
        }
    }

    /**
     * hssf读取表格文件，读取内容过程通过此对象回调出来
     * 目标解析结果：表格中每行数据拼接成一个List
     *
     * @see #parseXlsx(File, IReadCallback)
     */
    private final class XlsListener implements HSSFListener {

        private FormatTrackingHSSFListener formatListener;

        private final IReadCallback callback;

        /** 一行数据，每个单元格数据以逗号分隔 */
        private List<String> rowInfo;

        private SSTRecord sstRecord;

        private XlsListener(@NonNull IReadCallback callback) {
            this.callback = callback;
            MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
            formatListener = new FormatTrackingHSSFListener(listener);
        }

        @Override
        public void processRecord(Record record) {
            String cell = null;
            switch (record.getSid()) {
                case BOFRecord.sid:
                    break;
                case SSTRecord.sid:
                    sstRecord = (SSTRecord) record;
                    break;
                case LabelSSTRecord.sid:
                    LabelSSTRecord lsrec = (LabelSSTRecord) record;
                    if (sstRecord != null) {
                        cell = sstRecord.getString(lsrec.getSSTIndex()).toString();
                    }
                    break;
                case LabelRecord.sid:
                    cell = ((LabelRecord) record).getValue();
                    break;
                case NumberRecord.sid:
                    //LYX_TAG 2020/5/8 14:50 日期格式的单元格内容也可能会被转成这个NumberRecord，而不是字符串
                    //因此日期和数字会混在一起无法分辨
                    NumberRecord nr = (NumberRecord)record;
                    cell = nr.getValue() + "";
                    break;
                case BoundSheetRecord.sid:
                    break;
                case FormulaRecord.sid:
                    //被转换成了double值
                    cell = ((FormulaRecord)record).getValue() + "";
                    break;
                case StringRecord.sid:
                    cell = ((StringRecord)record).getString();
                    break;
                case BlankRecord.sid:
                    cell = "";
                    break;
                case BoolErrRecord.sid:
                    cell = ((BoolErrRecord)record).getBooleanValue() + "";
                    break;
                case NoteRecord.sid:
                    break;
                case RKRecord.sid:
                    break;
                default:
                    break;
            }

            if (rowInfo == null) {
                rowInfo = new ArrayList<>();
            }

            //只取单元格的内容
            if (cell != null && (record instanceof CellRecord)) {
                rowInfo.add(cell);
            }

            boolean isLastCellOfRow = record instanceof LastCellOfRowDummyRecord;
            if (isLastCellOfRow) {
                //行结束，给外面回调
                callback.onRowRead(((LastCellOfRowDummyRecord)record).getRow(), rowInfo);
                rowInfo.clear();
                rowInfo = null;
            }
        }
    }
}
