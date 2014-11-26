package testExport;

public class Book {
	private int bookId;
	private String name;
	private String author;
	private float price;
	private String isbn;
	private String pubName;
	private byte[] preface;

	public Book() {
	}

	public Book(int bookId, String name, String author, float price,
			String isbn, String pubName, byte[] preface) {
		this.bookId = bookId;
		this.name = name;
		this.author = author;
		this.price = price;
		this.isbn = isbn;
		this.pubName = pubName;
		this.preface = preface;
	}

	public int getBookId() {
		return bookId;
	}

	public void setBookId(int bookId) {
		this.bookId = bookId;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public float getPrice() {
		return price;
	}

	public void setPrice(float price) {
		this.price = price;
	}

	public String getIsbn() {
		return isbn;
	}

	public void setIsbn(String isbn) {
		this.isbn = isbn;
	}

	public String getPubName() {
		return pubName;
	}

	public void setPubName(String pubName) {
		this.pubName = pubName;
	}

	public byte[] getPreface() {
		return preface;
	}

	public void setPreface(byte[] preface) {
		this.preface = preface;
	}
}