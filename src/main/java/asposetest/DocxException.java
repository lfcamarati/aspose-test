package asposetest;

public class DocxException extends RuntimeException {

	public DocxException(String message) {
		super(message);
	}

	public DocxException(Throwable throwable) {
		super(throwable);
	}

	public DocxException(String message, Throwable cause) {
		super(message, cause);
	}
}
