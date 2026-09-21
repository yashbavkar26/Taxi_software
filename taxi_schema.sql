CREATE DATABASE IF NOT EXISTS taxi_database;
USE taxi_database;

CREATE TABLE IF NOT EXISTS drivers (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    driver VARCHAR(255) UNIQUE,
    contact VARCHAR(20)
);

CREATE TABLE IF NOT EXISTS bookings (
    id INT AUTO_INCREMENT PRIMARY KEY,
    received BOOLEAN DEFAULT FALSE,
    sl_no INT,
    coordinator VARCHAR(255),
    guest_name VARCHAR(255),
    guest_mobile_no VARCHAR(20),
    travel_date DATE,
    pickup_location VARCHAR(255),
    drop_location VARCHAR(255),
    start_time TIME,
    return_time TIME,
    reason TEXT,
    driver_name VARCHAR(255),
    contact VARCHAR(20),
    vehicle_type VARCHAR(255),
    driver_message TEXT GENERATED ALWAYS AS (
        CONCAT('Tomorrow', CHAR(10), COALESCE(guest_name, ''), ' ', COALESCE(guest_mobile_no, ''), CHAR(10), COALESCE(pickup_location, ''), ' to ', COALESCE(drop_location, ''), CHAR(10), COALESCE(TIME_FORMAT(start_time, '%l:%i %p'), ''))
    ) STORED,
    guest_message TEXT GENERATED ALWAYS AS (
        CONCAT('From GIM', CHAR(10), 'Driver ', COALESCE(driver_name, ''), ' ', COALESCE(contact, ''), ' will come to pick you tomorrow at ', COALESCE(TIME_FORMAT(start_time, '%l:%i %p'), ''), ' at ', COALESCE(pickup_location, ''))
    ) STORED,
    FOREIGN KEY (driver_name) REFERENCES drivers(driver)
);

CREATE TABLE IF NOT EXISTS masterfile (
    id INT AUTO_INCREMENT PRIMARY KEY,
    sl_no INT,
    vendor VARCHAR(255),actoractor_id
    travel_date DATE,
    vehicle_no VARCHAR(50),
    vehicle_type VARCHAR(100),
    starting_km DECIMAL(10, 2),
    ending_km DECIMAL(10, 2),
    total_km DECIMAL(10, 2) GENERATED ALWAYS AS (ending_km - starting_km) STORED,
    night DECIMAL(10, 2) DEFAULT 0.00,
    rate DECIMAL(10, 2) DEFAULT 0.00,
    toll DECIMAL(10, 2) DEFAULT 0.00,
    amt DECIMAL(10, 2) DEFAULT 0.00,
    guest_name VARCHAR(255),
    total_amt DECIMAL(10, 2) GENERATED ALWAYS AS (COALESCE(amt, 0) + COALESCE(toll, 0) + COALESCE(night, 0)) STORED,
    in_time TIME,
    out_time TIME,
    driver_name VARCHAR(255),
    from_to VARCHAR(255)
);

-- --------------------------------------------------------
-- Vendor Annexure View for 'Arvind'
-- Automatically reflects any new rows added to masterfile
-- --------------------------------------------------------
CREATE OR REPLACE VIEW view_annexure_arvind AS
SELECT 
    ROW_NUMBER() OVER (ORDER BY travel_date, id) AS sl_no,
    travel_date,
    guest_name,
    vehicle_no,
    vehicle_type,
    starting_km AS starting_reading,
    ending_km AS ending_reading,
    total_km AS total_kms,
    rate AS rate_per_km_above_80,
    amt AS total_km_rs,
    toll AS toll_charges,
    night AS night_charges,
    0.00 AS waiting_charges,
    total_amt AS total_rs,
    CONCAT_WS(' - ', TIME_FORMAT(in_time, '%l:%i %p'), TIME_FORMAT(out_time, '%l:%i %p')) AS timing,
    driver_name AS driver
FROM masterfile
WHERE vendor = 'Arvind';

-- --------------------------------------------------------
-- Stored Procedure: Create a new annexure view dynamically 
-- whenever a new vendor is added.
-- Usage example: CALL CreateVendorAnnexureView('Rahul');
-- --------------------------------------------------------
DELIMITER //

CREATE PROCEDURE CreateVendorAnnexureView(IN p_vendor VARCHAR(255))
BEGIN
    SET @clean_name = LOWER(REPLACE(p_vendor, ' ', '_'));
    SET @sql = CONCAT(
        'CREATE OR REPLACE VIEW view_annexure_', @clean_name, ' AS ',
        'SELECT ',
        '    ROW_NUMBER() OVER (ORDER BY travel_date, id) AS sl_no, ',
        '    travel_date, ',
        '    guest_name, ',
        '    vehicle_no, ',
        '    vehicle_type, ',
        '    starting_km AS starting_reading, ',
        '    ending_km AS ending_reading, ',
        '    total_km AS total_kms, ',
        '    rate AS rate_per_km_above_80, ',
        '    amt AS total_km_rs, ',
        '    toll AS toll_charges, ',
        '    night AS night_charges, ',
        '    0.00 AS waiting_charges, ',
        '    total_amt AS total_rs, ',
        '    CONCAT_WS(\' - \', TIME_FORMAT(in_time, \'%l:%i %p\'), TIME_FORMAT(out_time, \'%l:%i %p\')) AS timing, ',
        '    driver_name AS driver ',
        'FROM masterfile ',
        'WHERE vendor = \'', p_vendor, '\';'
    );
    PREPARE stmt FROM @sql;
    EXECUTE stmt;
    DEALLOCATE PREPARE stmt;
END //

DELIMITER ;


USE taxi_database

