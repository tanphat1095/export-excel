package vn.com.phat.excel.exception;

public class ExcelException extends Exception{

    /**
     * Constructs a new ExcelException by wrapping the specified exception.
     *
     * @param exception the underlying exception to be wrapped
     */
    public ExcelException(Exception exception) {
        super(exception);
    }
}
