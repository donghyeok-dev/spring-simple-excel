package com.github.donghyeok.excel.exception;

public class ExcelReaderException extends RuntimeException{

    public ExcelReaderException(String message) {
        super(message);
    }

    public ExcelReaderException(String message, Throwable cause) {
        super(message, cause);
    }
}