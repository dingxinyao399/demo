package com.tryinfo.utils;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.net.URL;
import java.util.Collection;
import java.util.regex.Pattern;

import lombok.NonNull;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
/**
 * ExcelPoiUtils class
 *
 * @author dingxinyao
 * @date 2019/6/4
 */
public class ExcelPoiUtils {
	private Workbook workbook;
	private Sheet sheet;
	private Row row;
	private Cell cell;
	private String cellValue;
	public String suffix;
	private CellStyle cellStyle;
    private CellStyle contextStyle;
	 DataFormat df;
	public void initExcel(String excelPath,int sheetIndex) {
		File file = new File(excelPath);
		if(!file.exists()) {
			//文件不存在
			workbook = null;
			sheet = null;
		}else {
			try {
				FileInputStream inputStream = new FileInputStream(file);
				if(excelPath.endsWith(".xls")) {
					POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
					workbook = new HSSFWorkbook(poifsFileSystem);
					this.suffix = ".xls";
					
					sheet = workbook.getSheetAt(sheetIndex);
					cellStyle = workbook.createCellStyle();
                    contextStyle = workbook.createCellStyle();
                    df = workbook.createDataFormat();
                    cellStyle.setLocked(false);
					cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
					inputStream.close();
				}else  if(excelPath.endsWith(".xlsx")){
					workbook = new XSSFWorkbook(inputStream);
					this.suffix = ".xlsx";
					
					sheet = workbook.getSheetAt(sheetIndex);
					cellStyle = workbook.createCellStyle();
					cellStyle.setLocked(false);
                    contextStyle = workbook.createCellStyle();
                    df = workbook.createDataFormat();
					cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
					inputStream.close();
				}
                //设置边框样式
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                //设置边框颜色z
                cellStyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
                cellStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
                cellStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
                cellStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
                //设置边框样式
                contextStyle.setBorderTop(BorderStyle.THIN);
                contextStyle.setBorderBottom(BorderStyle.THIN);
                contextStyle.setBorderLeft(BorderStyle.THIN);
                contextStyle.setBorderRight(BorderStyle.THIN);
                //设置边框颜色
                contextStyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
                contextStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
                contextStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
                contextStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				workbook = null;
				sheet = null;
			}
		}
	}

    public void switchExcel(int sheetIndex) {
        sheet = workbook.getSheetAt(sheetIndex);
        sheet.setForceFormulaRecalculation(true);
    }
	public void initURLExcel(String excelPath,int sheetIndex) {
		try {
			URL url = new URL(excelPath);
			BufferedInputStream inputStream = new BufferedInputStream(url.openStream(),1024*1024*10);
			if(excelPath.endsWith(".xls")) {
				POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream); 
				workbook = new HSSFWorkbook(poifsFileSystem);
				this.suffix = ".xls";
				
				sheet = workbook.getSheetAt(sheetIndex);
				cellStyle = workbook.createCellStyle();
				cellStyle.setLocked(false);
				cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
				inputStream.close();
			}else if(excelPath.endsWith(".xlsx")){
				workbook = new XSSFWorkbook(inputStream);
				this.suffix = ".xlsx";
				
				sheet = workbook.getSheetAt(sheetIndex);
				cellStyle = workbook.createCellStyle();
				cellStyle.setLocked(false);
				cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
				inputStream.close();
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			workbook = null;
			sheet = null;
		}
	}

    public void workbookFactoryURLExcel(String excelPath, int sheetIndex) {
        try {
            URL url = new URL(excelPath);
            BufferedInputStream inputStream = new BufferedInputStream(url.openStream(), 1024 * 1024 * 10);
            if (excelPath.endsWith(".xls")) {
                POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
                workbook = WorkbookFactory.create(poifsFileSystem);
                this.suffix = ".xls";
            } else {
                workbook = WorkbookFactory.create(inputStream);
                this.suffix = ".xlsx";
            }
            sheet = workbook.getSheetAt(sheetIndex);
            cellStyle = workbook.createCellStyle();
            cellStyle.setLocked(false);
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
            inputStream.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
        	if(null != workbook) {
        		try {
					workbook.close();
				} catch (IOException e1) {
        		    e1.printStackTrace();
				}
        	}
        	
            workbook = null;
            sheet = null;
            
            e.printStackTrace();
        }
    }
	
	public int getAllRows() {
		if(null != sheet) {
			int rows = sheet.getLastRowNum()+1;
			return rows;
		}
		return 0;
	}
	
	public String read(int rowIndex,int colIndex) {
		cellValue = "";
		if(null != sheet) {
			row = sheet.getRow(rowIndex);
			if(null != row) {
				cell = row.getCell(colIndex);
				if(null != cell) {
					cell.setCellType(CellType.STRING);
					cellValue = cell.getStringCellValue();
				}
			}
		}
		return cellValue;
	}
	
	public void deleteRow(int rowIndex) {
		if(null != sheet) {
			int lastRowNum=sheet.getLastRowNum();
		    if(rowIndex>=0&&rowIndex<lastRowNum) {
                //将行号为rowIndex+1一直到行号为lastRowNum的单元格全部上移一行，以便删除rowIndex行
                sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
            }
		    if(rowIndex==lastRowNum){
		    	row=sheet.getRow(rowIndex);
		        if(row!=null)
		            sheet.removeRow(row);
		    }

	
		}
	}

    public Row getCell(int rowIndex) {
        if (null != sheet) {
            row = sheet.getRow(rowIndex);
        }
        return row;
    }

    public void writeCell(Row row, int rowIndex, String msg) {
        if (null != sheet) {
        	if(null == row) {
        		row = sheet.createRow(rowIndex);
        	}
        	Cell cell2 = row.getCell(18);
            if (cell2 == null) {
                //创建最后一列
                cell2 = row.createCell(18);
            }
            Font ztFont = workbook.createFont();
            // 将字体设置为“红色”
            ztFont.setColor(Font.COLOR_RED);
            // 将字体大小设置为18px
            ztFont.setFontHeightInPoints((short) 12);
            // 将“华文行楷”字体应用到当前单元格上
            ztFont.setFontName("华文行楷");
            CellStyle style = workbook.createCellStyle();
            style.setFont(ztFont);
            cell2.setCellType(CellType.STRING);
            cell2.setCellStyle(style);
            cell2.setCellValue(msg);
        }
    }
    
    
    public void setCellRed(Row selectedRow, int rowIndex) {
        if (null != sheet) {
        
        	
        	Cell cell2 = selectedRow.getCell(rowIndex);
        	if(cell2==null) {
        		cell2 = row.createCell(rowIndex);
        	}
        	    
            CellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(IndexedColors.RED.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell2.setCellType(CellType.STRING);
            cell2.setCellStyle(style);
          
        }
    }

    public void closeExcel() {
    	if(null != workbook) {
    		try {
    			workbook.close();
    		} catch (Exception e) {
    		}
    	}
	}
	
	public void write(int rowIndex,int colIndex,String value) {
		if(null != sheet) {
			row = sheet.getRow(rowIndex);
			if(null == row) {
				row = sheet.createRow(rowIndex);
			}
			cell = row.getCell(colIndex);
			if(null == cell) {
				cell = row.createCell(colIndex);
			}
			cell.setCellType(CellType.STRING);
			cell.setCellStyle(cellStyle);
            if (StringUtils.isBlank(value)) {
                value = "";
            }
			cell.setCellValue(value+"");
		}
	}

    public void writeStyle(int rowIndex, int colIndex, String value) {
        if (null != sheet) {
            row = sheet.getRow(rowIndex);
            if (null == row) {
                row = sheet.createRow(rowIndex);
            }
            cell = row.getCell(colIndex);
            if (null == cell) {
                cell = row.createCell(colIndex);
            }

            cell.setCellType(CellType.STRING);
            cell.setCellStyle(contextStyle);
            if (StringUtils.isBlank(value)) {
                value = "";
            }
            cell.setCellValue(value + "");
        }
    }

    /**
     * 通过注解反射对List<Entity>中的需要累加的字段进行累加
     *
     * @param list List<Entity>
     * @param t    需要反射的实体
     * @param <E>
     * @return E
     */
    public <E> E listEntityAdd(@NonNull Collection<E> list, E t) {
        Field[] fields = FieldUtils.getFieldsWithAnnotation(t.getClass(), Add.class);
        for (E e : list) {
            for (Field field : fields) {
                String fieldName = field.getName();
                try {
                    int value = (int) FieldUtils.readDeclaredField(e, fieldName, true);
                    int valueTotal = FieldUtils.readDeclaredField(t, fieldName, true) == null ? 0 : (int) FieldUtils.readDeclaredField(t, fieldName, true);
                    FieldUtils.writeDeclaredField(t, fieldName, value + valueTotal, true);
                } catch (Exception e1) {
                    e1.printStackTrace();
                }

            }
        }
        return t;
    }

    /**
     * 通过注解反射对List<Entity>中的需要累加的字段进行累加
     *
     * @param list List<Entity>
     * @param t    需要反射的实体
     * @param <E>
     * @return E
     */
    public <E> E listEntityAddByDouble(@NonNull Collection<E> list, E t) {
        Field[] fields = FieldUtils.getFieldsWithAnnotation(t.getClass(), AddDouble.class);
        for (E e : list) {
            for (Field field : fields) {
                String fieldName = field.getName();
                Class<?> type = field.getType();
                AddDouble annotation = field.getAnnotation(AddDouble.class);
                try {
                    AddHandler<? extends Number> addHandler = annotation.handler().getConstructor().newInstance();
                    if (Number.class.isAssignableFrom(type)) {
                        Number value = (Number) FieldUtils.readDeclaredField(e, fieldName, true);
                        Number valueTotal = FieldUtils.readDeclaredField(t, fieldName, true) == null ? 0 : (Number) FieldUtils.readDeclaredField(t, fieldName, true);
                        Number addResult = addHandler.add(value, valueTotal);
                        FieldUtils.writeDeclaredField(t, fieldName, addHandler.get(addResult), true);
                    }
                } catch (Exception e1) {
                    e1.printStackTrace();
                }
            }
        }
        Field[] integerFields = FieldUtils.getFieldsWithAnnotation(t.getClass(), Add.class);
        for (E e : list) {
            for (Field field : integerFields) {
                String fieldName = field.getName();
                try {
                    int value = (int) FieldUtils.readDeclaredField(e, fieldName, true);
                    int valueTotal = FieldUtils.readDeclaredField(t, fieldName, true) == null ? 0 : (int) FieldUtils.readDeclaredField(t, fieldName, true);
                    FieldUtils.writeDeclaredField(t, fieldName, value + valueTotal, true);
                } catch (Exception e1) {
                    e1.printStackTrace();
                }
            }
        }
        return t;
    }


    public void statisticsWrite(int rowIndex, int colIndex, String value) {
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, 0, 7);
        sheet.addMergedRegion(region);
        // 下边框
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        // 左边框
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        // 有边框
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
        // 上边框
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);

        if (null != sheet) {
            row = sheet.getRow(rowIndex);
            if (null == row) {
                row = sheet.createRow(rowIndex);
            }
            row.setHeightInPoints(30);
            cell = row.getCell(colIndex);
            if (null == cell) {
                cell = row.createCell(colIndex);
            }
            Font ztFont = workbook.createFont();
            // 将字体大小设置为13px
            ztFont.setFontHeightInPoints((short) 14);
            // 将“华文行楷”字体应用到当前单元格上
            ztFont.setFontName("仿宋_GB2312");
            cellStyle.setFont(ztFont);
            cell.setCellType(CellType.STRING);
            cell.setCellStyle(cellStyle);
            if (StringUtils.isBlank(value)) {
                value = "";
            }
            cell.setCellValue(value + "");
        }
    }

    private static Pattern NUMBER_PATTERN = Pattern.compile("^[-\\+]?[\\d]*$");
    /**
     * 判断是否为整数
     * @param str 传入的字符串
     * @return 是整数返回true,否则返回false
     */
    public  boolean isInteger(String str) {
        return NUMBER_PATTERN.matcher(str).matches();
    }
    public void writeInteger(int rowIndex,int colIndex,String value) {
        //数据格式只显示整数
          contextStyle.setDataFormat(df.getFormat("0"));
          int num ;
          if(StringUtils.isBlank(value)){
              num=0;
          }else{
              num=Integer.parseInt(value);
          }
          if(null != sheet) {
              row = sheet.getRow(rowIndex);
              if(null == row) {
                  row = sheet.createRow(rowIndex);
              }
              cell = row.getCell(colIndex);
              if(null == cell) {
                  cell = row.createCell(colIndex);
              }
              cell.setCellStyle(contextStyle);
              cell.setCellValue(num);
          }
        }
       
    public void writerNum(int rowIndex,int colIndex,String value) {
    	double num ;
        if(StringUtils.isBlank(value)){
            num=0;
        }else{
            num=Double.valueOf(value);
        }
        if(null != sheet) {
            row = sheet.getRow(rowIndex);
            if(null == row) {
                row = sheet.createRow(rowIndex);
            }
            cell = row.getCell(colIndex);
            if(null == cell) {
                cell = row.createCell(colIndex);
            }
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(num);
        }
       
    }
	
	public void save(String excelPath) {
		if(null != workbook) {
			FileOutputStream fileoutputStream = null;
			try {
				fileoutputStream = new FileOutputStream(new File(excelPath));
	            workbook.write(fileoutputStream);
	        } catch (Exception e) {
	            e.printStackTrace();
	        }finally {
	        	try {
	        		fileoutputStream.flush();
				} catch (Exception e2) {
				}
	        	try {
	        		fileoutputStream.close();
				} catch (Exception e2) {
				}
	        }
		}
	}


    /**
     *
     * @param request
     * @param response
     * @param url 文件地址
     * @param fileName 文件名
     * @throws Exception
     */
	public void download(HttpServletRequest request, HttpServletResponse response,String url,String fileName)throws  Exception{
        response.setHeader("content-disposition",
                "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));
        // 读取要下载的文件，保存到文件输入流
        FileInputStream in = new FileInputStream(url);
        // 创建输出流
        OutputStream out = response.getOutputStream();
        // 创建缓冲区
        byte[] buffer = new byte[1024];
        int len = 0;
        // 循环将输入流中的内容读取到缓冲区中
        while ((len = in.read(buffer)) > 0) {
            // 输出缓冲区内容到浏览器，实现文件下载
            out.write(buffer, 0, len);
        }
        // 关闭文件流
        in.close();
        // 关闭输出流
        out.close();
    }
}
