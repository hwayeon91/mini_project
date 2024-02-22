CREATE TABLE users (
    id NUMBER PRIMARY KEY,
    name VARCHAR2(100) NOT NULL,
    username VARCHAR2(50) UNIQUE NOT NULL,
    password VARCHAR2(100) NOT NULL,
    phone_number VARCHAR2(20),
    address VARCHAR2(200),
    user_level VARCHAR2(20) CHECK (user_level IN ('관리자', 'VIP', 'Gold', 'Silver', 'Bronze'))
);

CREATE TABLE products (
    product_code VARCHAR2(50) PRIMARY KEY,
    product_title VARCHAR2(200) NOT NULL, -- 작품 제목 추가
    author VARCHAR2(100),
    publisher VARCHAR2(100),
    publish_date DATE,
    price NUMBER,
    subject VARCHAR2(100),
    grade VARCHAR2(50)
);

CREATE TABLE reviews (
    review_id NUMBER PRIMARY KEY,
    user_id NUMBER,
    product_code VARCHAR2(50), -- 상품 코드 추가
    title VARCHAR2(200) NOT NULL,
    content VARCHAR2(1000) NOT NULL,
    views NUMBER DEFAULT 0,
    FOREIGN KEY (user_id) REFERENCES users(id),
    FOREIGN KEY (product_code) REFERENCES products(product_code) -- 외래 키 추가
);
