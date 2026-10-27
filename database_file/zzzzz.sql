-- --------------------------------------------------------
-- Host:                         127.0.0.1
-- Server version:               8.2.0 - MySQL Community Server - GPL
-- Server OS:                    Win64
-- HeidiSQL Version:             12.6.0.6765
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!50503 SET NAMES utf8mb4 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;


-- Dumping database structure for attendance_management_db
DROP DATABASE IF EXISTS `attendance_management_db`;
CREATE DATABASE IF NOT EXISTS `attendance_management_db` /*!40100 DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci */ /*!80016 DEFAULT ENCRYPTION='N' */;
USE `attendance_management_db`;

-- Dumping structure for table attendance_management_db.attendance
DROP TABLE IF EXISTS `attendance`;
CREATE TABLE IF NOT EXISTS `attendance` (
  `id` int NOT NULL AUTO_INCREMENT,
  `student_id` int DEFAULT NULL,
  `attendance_date` date DEFAULT NULL,
  `status` varchar(10) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=39 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- Dumping data for table attendance_management_db.attendance: 32 rows
/*!40000 ALTER TABLE `attendance` DISABLE KEYS */;
REPLACE INTO `attendance` (`id`, `student_id`, `attendance_date`, `status`) VALUES
	(5, 12, '2026-10-13', 'Present'),
	(6, 13, '2026-10-13', 'Present'),
	(7, 11, '2026-10-13', 'Present'),
	(8, 10, '2026-10-13', 'Present'),
	(9, 9, '2026-10-13', 'Present'),
	(10, 8, '2026-10-13', 'Absent'),
	(11, 7, '2026-10-13', 'Absent'),
	(12, 7, '2026-10-13', 'Absent'),
	(13, 6, '2026-10-13', 'Absent'),
	(18, 8, '2026-10-13', 'Absent'),
	(17, 14, '2026-10-13', 'Present'),
	(16, 13, '2026-10-13', 'Present'),
	(19, 13, '2026-10-13', 'Absent'),
	(20, 11, '2026-10-15', 'Present'),
	(21, 11, '2026-10-15', 'Present'),
	(22, 12, '2026-10-15', 'Present'),
	(23, 14, '2026-10-15', 'Present'),
	(24, 6, '2026-10-15', 'Absent'),
	(25, 7, '2026-10-15', 'Absent'),
	(26, 8, '2026-10-15', 'Absent'),
	(27, 9, '2026-10-15', 'Absent'),
	(28, 10, '2026-10-15', 'Absent'),
	(29, 6, '2026-10-16', 'Absent'),
	(30, 7, '2026-10-16', 'Absent'),
	(31, 8, '2026-10-16', 'Absent'),
	(32, 9, '2026-10-16', 'Absent'),
	(33, 10, '2026-10-16', 'Absent'),
	(34, 11, '2026-10-16', 'Absent'),
	(35, 12, '2026-10-16', 'Absent'),
	(36, 13, '2026-10-16', 'Absent'),
	(37, 14, '2026-10-16', 'Absent'),
	(38, 15, '2026-10-16', 'Present');
/*!40000 ALTER TABLE `attendance` ENABLE KEYS */;

-- Dumping structure for table attendance_management_db.students
DROP TABLE IF EXISTS `students`;
CREATE TABLE IF NOT EXISTS `students` (
  `id` int NOT NULL AUTO_INCREMENT,
  `school_id` varchar(50) DEFAULT NULL,
  `student_name` varchar(50) DEFAULT NULL,
  `roll_no` varchar(50) DEFAULT NULL,
  `class_name` varchar(50) DEFAULT NULL,
  `contact_no` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=16 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- Dumping data for table attendance_management_db.students: 10 rows
/*!40000 ALTER TABLE `students` DISABLE KEYS */;
REPLACE INTO `students` (`id`, `school_id`, `student_name`, `roll_no`, `class_name`, `contact_no`) VALUES
	(14, 'S420', 'shoaib', '100', '12', NULL),
	(6, 'S420', 'Harry', '101', '12', NULL),
	(7, 'S420', 'Jack', '102', '11', NULL),
	(8, 'S420', 'Luther', '99', '10', NULL),
	(9, 'S420', 'Addy', '112', '9', NULL),
	(10, 'S420', 'Denny', '90', '8', NULL),
	(11, 'S420', 'Lukas', '80', '8', NULL),
	(12, 'S420', 'Naruto', '70', '9', NULL),
	(13, 'S420', 'Gojo', '95', '10', NULL),
	(15, 'S420', 'Hulk', '115', '12', NULL);
/*!40000 ALTER TABLE `students` ENABLE KEYS */;

/*!40103 SET TIME_ZONE=IFNULL(@OLD_TIME_ZONE, 'system') */;
/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IFNULL(@OLD_FOREIGN_KEY_CHECKS, 1) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40111 SET SQL_NOTES=IFNULL(@OLD_SQL_NOTES, 1) */;
