package com.github.donghyeok.excel.exception;

public class ExcelWriterException extends RuntimeException{

    public ExcelWriterException(String message) {
        super(message);
    }

    public ExcelWriterException(String message, Throwable cause) {
        super(message, cause);
    }
}