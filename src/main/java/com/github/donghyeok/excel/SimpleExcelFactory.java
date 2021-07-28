package com.github.donghyeok.excel;

class SimpleExcelFactory {
    public <T> SimpleExcelReader<T> createExcelReader(Class<T> tClass) {
        return new SimpleExcelReader<>(tClass);
    }
}
