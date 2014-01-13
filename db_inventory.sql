-- phpMyAdmin SQL Dump
-- version 4.0.4
-- http://www.phpmyadmin.net
--
-- Host: localhost
-- Generation Time: Dec 24, 2013 at 12:57 AM
-- Server version: 5.6.12-log
-- PHP Version: 5.4.16

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- Database: `db_inventory`
--
CREATE DATABASE IF NOT EXISTS `db_inventory` DEFAULT CHARACTER SET latin1 COLLATE latin1_swedish_ci;
USE `db_inventory`;

-- --------------------------------------------------------

--
-- Table structure for table `tbl_category`
--

CREATE TABLE IF NOT EXISTS `tbl_category` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Category_Code` varchar(200) NOT NULL,
  `Category_Name` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=3 ;

--
-- Dumping data for table `tbl_category`
--

INSERT INTO `tbl_category` (`No`, `Category_Code`, `Category_Name`) VALUES
(1, 'Sof', 'Sofa'),
(2, 'DinTab', 'Dining Table');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_company`
--

CREATE TABLE IF NOT EXISTS `tbl_company` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `Name` varchar(200) NOT NULL,
  `Owner` varchar(200) NOT NULL,
  `Mobile_Number` varchar(100) NOT NULL,
  `Address` varchar(200) NOT NULL,
  `Description` varchar(200) NOT NULL,
  `Mission` varchar(200) NOT NULL,
  `Vision` varchar(200) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=2 ;

--
-- Dumping data for table `tbl_company`
--

INSERT INTO `tbl_company` (`ID`, `Name`, `Owner`, `Mobile_Number`, `Address`, `Description`, `Mission`, `Vision`) VALUES
(1, 'Caparal Furniture Shop', 'sample', '0911111111', 'sample', 'sample', 'sample', 'sample');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_customer`
--

CREATE TABLE IF NOT EXISTS `tbl_customer` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Customer_ID` varchar(200) NOT NULL,
  `Customer_Name` varchar(200) NOT NULL,
  `Representative1` varchar(200) NOT NULL,
  `Representative2` varchar(200) NOT NULL,
  `Mobile_Number` varchar(200) NOT NULL,
  `Address` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=3 ;

--
-- Dumping data for table `tbl_customer`
--

INSERT INTO `tbl_customer` (`No`, `Customer_ID`, `Customer_Name`, `Representative1`, `Representative2`, `Mobile_Number`, `Address`) VALUES
(1, 'C-00001', 'Home Plus', 'Pedro Magtibay', '', '091111111', 'Cavite City'),
(2, 'C-00002', 'Perlin Diloy', 'Perlin Diloy', '', '09111111', 'Cavite City');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_logs`
--

CREATE TABLE IF NOT EXISTS `tbl_logs` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Username` varchar(100) NOT NULL,
  `Login` varchar(100) NOT NULL,
  `Logout` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=288 ;

--
-- Dumping data for table `tbl_logs`
--

INSERT INTO `tbl_logs` (`No`, `Username`, `Login`, `Logout`) VALUES
(1, 'admin', '12/5/2013 10:02:00 PM', '12/5/2013 10:15:07 PM'),
(2, 'admin', '12/5/2013 10:16:13 PM', '12/5/2013 10:16:15 PM'),
(3, 'admin', '12/5/2013 10:17:01 PM', '12/5/2013 10:17:02 PM'),
(4, 'admin', '12/5/2013 10:17:19 PM', '12/5/2013 10:18:16 PM'),
(5, 'admin', '12/5/2013 10:18:15 PM', '12/5/2013 10:18:16 PM'),
(6, 'admin', '12/5/2013 10:18:32 PM', '12/5/2013 10:57:17 PM'),
(7, 'admin', '12/5/2013 10:20:13 PM', '12/5/2013 10:57:17 PM'),
(8, 'admin', '12/5/2013 10:21:03 PM', '12/5/2013 10:57:17 PM'),
(9, 'admin', '12/5/2013 10:34:28 PM', '12/5/2013 10:57:17 PM'),
(10, 'admin', '12/5/2013 11:04:19 PM', '12/5/2013 11:04:57 PM'),
(11, 'admin', '12/5/2013 11:04:40 PM', '12/5/2013 11:04:57 PM'),
(12, 'admin', '12/5/2013 11:09:52 PM', '12/5/2013 11:10:28 PM'),
(13, 'admin', '12/5/2013 11:54:43 PM', '12/5/2013 11:54:54 PM'),
(14, 'admin', '12/5/2013 11:55:40 PM', '12/5/2013 11:59:25 PM'),
(15, 'admin', '12/5/2013 11:57:12 PM', '12/5/2013 11:59:25 PM'),
(16, 'admin', '12/5/2013 11:58:14 PM', '12/5/2013 11:59:25 PM'),
(17, 'admin', '12/5/2013 11:59:44 PM', '12/6/2013 12:00:21 AM'),
(18, 'admin', '12/6/2013 12:03:02 AM', '12/6/2013 12:03:30 AM'),
(19, 'admin', '12/6/2013 12:06:21 AM', '12/6/2013 12:06:26 AM'),
(20, 'julie', '12/6/2013 12:06:31 AM', '12/6/2013 12:06:34 AM'),
(21, 'admin', '12/6/2013 12:16:58 AM', '12/6/2013 12:32:46 AM'),
(22, 'admin', '12/6/2013 12:39:37 AM', '12/6/2013 12:39:50 AM'),
(23, 'admin', '12/6/2013 12:49:23 AM', '12/6/2013 12:49:33 AM'),
(24, 'admin', '12/6/2013 12:50:52 AM', '12/6/2013 12:53:10 AM'),
(25, 'admin', '12/6/2013 12:56:26 AM', '12/6/2013 12:56:34 AM'),
(26, 'admin', '12/6/2013 12:57:33 AM', '12/6/2013 12:59:00 AM'),
(27, 'admin', '12/6/2013 1:00:25 AM', '12/6/2013 1:00:58 AM'),
(28, 'admin', '12/6/2013 1:01:34 AM', '12/6/2013 1:01:53 AM'),
(29, 'admin', '12/6/2013 1:27:58 AM', '12/6/2013 1:29:20 AM'),
(30, 'admin', '12/6/2013 1:28:53 AM', '12/6/2013 1:29:20 AM'),
(31, 'admin', '12/6/2013 1:39:30 AM', '12/6/2013 1:39:53 AM'),
(32, 'admin', '12/6/2013 1:47:10 AM', '12/6/2013 1:47:25 AM'),
(33, 'admin', '12/6/2013 1:48:00 AM', '12/6/2013 1:48:59 AM'),
(34, 'admin', '12/6/2013 1:52:59 AM', '12/6/2013 1:54:09 AM'),
(35, 'admin', '12/6/2013 1:55:14 AM', '12/6/2013 1:56:41 AM'),
(36, 'admin', '12/6/2013 1:59:13 AM', '12/6/2013 1:59:32 AM'),
(37, 'admin', '12/6/2013 2:07:43 AM', '12/6/2013 2:08:04 AM'),
(38, 'admin', '12/6/2013 2:10:52 AM', '12/6/2013 2:11:08 AM'),
(39, 'admin', '12/6/2013 2:11:35 AM', '12/6/2013 2:12:09 AM'),
(40, 'admin', '12/6/2013 2:13:24 AM', '12/6/2013 2:13:42 AM'),
(41, 'admin', '12/6/2013 2:13:56 AM', '12/6/2013 2:18:48 AM'),
(42, 'admin', '12/6/2013 2:15:02 AM', '12/6/2013 2:18:48 AM'),
(43, 'admin', '12/6/2013 2:16:34 AM', '12/6/2013 2:18:48 AM'),
(44, 'admin', '12/6/2013 2:18:09 AM', '12/6/2013 2:18:48 AM'),
(45, 'admin', '12/6/2013 2:26:36 AM', '12/6/2013 2:29:55 AM'),
(46, 'admin', '12/6/2013 2:27:32 AM', '12/6/2013 2:29:55 AM'),
(47, 'admin', '12/6/2013 2:28:16 AM', '12/6/2013 2:29:55 AM'),
(48, 'admin', '12/6/2013 2:34:37 AM', '12/6/2013 2:35:28 AM'),
(49, 'admin', '12/6/2013 2:43:48 AM', '12/6/2013 2:44:01 AM'),
(50, 'admin', '12/6/2013 2:47:36 AM', '12/6/2013 2:47:56 AM'),
(51, 'admin', '12/6/2013 2:50:32 AM', '12/6/2013 2:55:17 AM'),
(52, 'admin', '12/6/2013 2:54:46 AM', '12/6/2013 2:55:17 AM'),
(53, 'admin', '12/6/2013 2:59:55 AM', '12/6/2013 3:00:27 AM'),
(54, 'admin', '12/6/2013 3:01:17 AM', '12/6/2013 3:01:47 AM'),
(55, 'admin', '12/9/2013 12:28:58 PM', '12/9/2013 12:32:35 PM'),
(56, 'admin', '12/9/2013 12:33:20 PM', '12/9/2013 12:33:23 PM'),
(57, 'admin', '12/9/2013 12:39:09 PM', '12/9/2013 12:39:22 PM'),
(58, 'admin', '12/9/2013 12:39:57 PM', '12/9/2013 12:40:04 PM'),
(59, 'admin', '12/9/2013 12:42:33 PM', '12/9/2013 12:45:19 PM'),
(60, 'admin', '12/9/2013 12:50:18 PM', '12/9/2013 1:06:06 PM'),
(61, 'admin', '12/9/2013 12:55:14 PM', '12/9/2013 1:06:06 PM'),
(62, 'admin', '12/9/2013 12:57:19 PM', '12/9/2013 1:06:06 PM'),
(63, 'admin', '12/9/2013 1:00:31 PM', '12/9/2013 1:06:06 PM'),
(64, 'admin', '12/9/2013 1:03:31 PM', '12/9/2013 1:06:06 PM'),
(65, 'admin', '12/9/2013 1:05:13 PM', '12/9/2013 1:06:06 PM'),
(66, 'admin', '12/9/2013 1:05:51 PM', '12/9/2013 1:06:06 PM'),
(67, 'admin', '12/9/2013 1:07:36 PM', '12/9/2013 1:11:18 PM'),
(68, 'admin', '12/9/2013 1:08:34 PM', '12/9/2013 1:11:18 PM'),
(69, 'admin', '12/9/2013 1:09:15 PM', '12/9/2013 1:11:18 PM'),
(70, 'admin', '12/9/2013 1:10:35 PM', '12/9/2013 1:11:18 PM'),
(71, 'admin', '12/9/2013 1:15:23 PM', '12/9/2013 1:19:24 PM'),
(72, 'admin', '12/9/2013 1:25:29 PM', '12/9/2013 1:26:03 PM'),
(73, 'admin', '12/9/2013 1:34:32 PM', '12/9/2013 1:41:50 PM'),
(74, 'admin', '12/9/2013 1:50:25 PM', '12/9/2013 1:51:25 PM'),
(75, 'admin', '12/9/2013 1:52:38 PM', '12/9/2013 1:52:46 PM'),
(76, 'admin', '12/9/2013 1:53:13 PM', '12/9/2013 1:53:35 PM'),
(77, 'admin', '12/9/2013 1:55:04 PM', '12/9/2013 1:57:40 PM'),
(78, 'admin', '12/9/2013 1:56:28 PM', '12/9/2013 1:57:40 PM'),
(79, 'admin', '12/9/2013 1:58:22 PM', '12/9/2013 2:00:20 PM'),
(80, 'admin', '12/9/2013 2:00:16 PM', '12/9/2013 2:00:20 PM'),
(81, 'admin', '12/9/2013 2:01:09 PM', '12/9/2013 2:01:27 PM'),
(82, 'admin', '12/9/2013 2:02:10 PM', '12/9/2013 2:09:47 PM'),
(83, 'admin', '12/9/2013 2:33:48 PM', '12/9/2013 2:40:57 PM'),
(84, 'admin', '12/9/2013 2:38:17 PM', '12/9/2013 2:40:57 PM'),
(85, 'admin', '12/9/2013 2:50:47 PM', '12/9/2013 2:51:20 PM'),
(86, 'admin', '12/9/2013 2:56:23 PM', '12/9/2013 3:13:28 PM'),
(87, 'admin', '12/9/2013 3:12:25 PM', '12/9/2013 3:13:28 PM'),
(88, 'admin', '12/9/2013 3:12:51 PM', '12/9/2013 3:13:28 PM'),
(89, 'admin', '12/9/2013 3:13:15 PM', '12/9/2013 3:13:28 PM'),
(90, 'admin', '12/9/2013 3:13:49 PM', '12/9/2013 3:16:46 PM'),
(91, 'admin', '12/9/2013 3:21:13 PM', '12/9/2013 3:22:27 PM'),
(92, 'admin', '12/9/2013 3:30:55 PM', '12/9/2013 3:31:30 PM'),
(93, 'admin', '12/9/2013 3:32:12 PM', '12/9/2013 3:32:26 PM'),
(94, 'admin', '12/9/2013 3:33:09 PM', '12/9/2013 3:33:20 PM'),
(95, 'admin', '12/9/2013 3:33:51 PM', '12/9/2013 3:34:06 PM'),
(96, 'admin', '12/9/2013 3:34:49 PM', '12/9/2013 3:35:11 PM'),
(97, 'admin', '12/9/2013 3:35:52 PM', '12/9/2013 3:37:07 PM'),
(98, 'admin', '12/9/2013 3:36:25 PM', '12/9/2013 3:37:07 PM'),
(99, 'admin', '12/9/2013 3:38:10 PM', '12/9/2013 3:39:10 PM'),
(100, 'admin', '12/9/2013 3:38:36 PM', '12/9/2013 3:39:10 PM'),
(101, 'admin', '12/9/2013 3:42:17 PM', '12/9/2013 3:42:29 PM'),
(102, 'admin', '12/9/2013 3:43:18 PM', '12/9/2013 3:45:08 PM'),
(103, 'admin', '12/9/2013 4:14:55 PM', '12/9/2013 4:17:22 PM'),
(104, 'admin', '12/9/2013 4:15:40 PM', '12/9/2013 4:17:22 PM'),
(105, 'admin', '12/9/2013 4:16:21 PM', '12/9/2013 4:17:22 PM'),
(106, 'admin', '12/9/2013 4:17:47 PM', '12/9/2013 4:18:31 PM'),
(107, 'admin', '12/9/2013 4:28:48 PM', '12/9/2013 4:29:36 PM'),
(108, 'admin', '12/9/2013 4:31:40 PM', '12/9/2013 4:32:11 PM'),
(109, 'admin', '12/9/2013 4:38:17 PM', '12/9/2013 4:38:24 PM'),
(110, 'admin', '12/9/2013 4:39:40 PM', '12/9/2013 4:39:45 PM'),
(111, 'admin', '12/9/2013 4:40:25 PM', '12/9/2013 4:40:38 PM'),
(112, 'admin', '12/9/2013 4:44:31 PM', '12/9/2013 4:47:58 PM'),
(113, 'admin', '12/9/2013 4:45:06 PM', '12/9/2013 4:47:58 PM'),
(114, 'admin', '12/9/2013 4:46:53 PM', '12/9/2013 4:47:58 PM'),
(115, 'admin', '12/9/2013 4:47:20 PM', '12/9/2013 4:47:58 PM'),
(116, 'admin', '12/9/2013 4:49:43 PM', '12/9/2013 4:49:54 PM'),
(117, 'admin', '12/9/2013 4:56:22 PM', '12/9/2013 4:57:26 PM'),
(118, 'admin', '12/9/2013 4:56:54 PM', '12/9/2013 4:57:26 PM'),
(119, 'admin', '12/9/2013 5:01:17 PM', '12/9/2013 5:02:03 PM'),
(120, 'admin', '12/9/2013 5:02:42 PM', '12/9/2013 5:02:55 PM'),
(121, 'admin', '12/9/2013 5:03:52 PM', '12/9/2013 5:03:59 PM'),
(122, 'admin', '12/9/2013 5:14:12 PM', '12/9/2013 5:17:33 PM'),
(123, 'admin', '12/9/2013 5:18:12 PM', '12/9/2013 5:23:41 PM'),
(124, 'admin', '12/9/2013 5:19:04 PM', '12/9/2013 5:23:41 PM'),
(125, 'admin', '12/9/2013 5:20:11 PM', '12/9/2013 5:23:41 PM'),
(126, 'admin', '12/9/2013 5:22:03 PM', '12/9/2013 5:23:41 PM'),
(127, 'admin', '12/9/2013 5:22:48 PM', '12/9/2013 5:23:41 PM'),
(128, 'admin', '12/9/2013 5:25:09 PM', '12/9/2013 5:25:38 PM'),
(129, 'admin', '12/9/2013 5:30:45 PM', '12/9/2013 5:31:23 PM'),
(130, 'admin', '12/9/2013 5:56:07 PM', '12/9/2013 5:56:47 PM'),
(131, 'admin', '12/9/2013 6:11:12 PM', '12/9/2013 6:12:14 PM'),
(132, 'admin', '12/9/2013 6:11:41 PM', '12/9/2013 6:12:14 PM'),
(133, 'admin', '12/9/2013 6:15:55 PM', '12/9/2013 6:16:39 PM'),
(134, 'admin', '12/9/2013 6:16:24 PM', '12/9/2013 6:16:39 PM'),
(135, 'admin', '12/9/2013 6:17:54 PM', '12/9/2013 6:19:12 PM'),
(136, 'admin', '12/9/2013 6:19:08 PM', '12/9/2013 6:19:12 PM'),
(137, 'admin', '12/9/2013 6:20:02 PM', '12/9/2013 6:22:21 PM'),
(138, 'admin', '12/9/2013 6:27:47 PM', '12/9/2013 6:27:57 PM'),
(139, 'admin', '12/9/2013 6:29:21 PM', '12/9/2013 6:29:32 PM'),
(140, 'admin', '12/9/2013 6:36:18 PM', '12/9/2013 6:36:25 PM'),
(141, 'admin', '12/9/2013 6:37:41 PM', '12/9/2013 6:37:54 PM'),
(142, 'admin', '12/9/2013 6:40:25 PM', '12/9/2013 6:40:33 PM'),
(143, 'admin', '12/9/2013 6:41:38 PM', '12/9/2013 6:41:50 PM'),
(144, 'admin', '12/9/2013 6:42:39 PM', '12/9/2013 6:42:59 PM'),
(145, 'admin', '12/9/2013 7:03:52 PM', '12/9/2013 7:04:24 PM'),
(146, 'admin', '12/9/2013 7:04:15 PM', '12/9/2013 7:04:24 PM'),
(147, 'admin', '12/9/2013 7:04:47 PM', '12/9/2013 7:05:15 PM'),
(148, 'admin', '12/9/2013 7:07:23 PM', '12/9/2013 7:07:40 PM'),
(149, 'admin', '12/9/2013 7:08:11 PM', '12/9/2013 7:08:21 PM'),
(150, 'admin', '12/9/2013 7:13:39 PM', '12/9/2013 7:13:47 PM'),
(151, 'admin', '12/9/2013 7:18:22 PM', '12/9/2013 7:19:26 PM'),
(152, 'admin', '12/9/2013 7:29:56 PM', '12/9/2013 7:30:13 PM'),
(153, 'admin', '12/9/2013 7:30:42 PM', '12/9/2013 7:31:00 PM'),
(154, 'admin', '12/9/2013 7:31:23 PM', '12/9/2013 7:31:43 PM'),
(155, 'admin', '12/9/2013 7:32:12 PM', '12/9/2013 7:32:41 PM'),
(156, 'admin', '12/9/2013 7:50:27 PM', '12/9/2013 7:50:45 PM'),
(157, 'admin', '12/9/2013 7:51:21 PM', '12/9/2013 7:52:30 PM'),
(158, 'admin', '12/9/2013 7:58:56 PM', '12/9/2013 7:59:05 PM'),
(159, 'admin', '12/9/2013 8:03:08 PM', '12/9/2013 8:04:00 PM'),
(160, 'admin', '12/9/2013 8:03:33 PM', '12/9/2013 8:04:00 PM'),
(161, 'admin', '12/9/2013 8:09:06 PM', '12/9/2013 8:09:20 PM'),
(162, 'admin', '12/9/2013 8:12:52 PM', '12/9/2013 8:13:03 PM'),
(163, 'admin', '12/9/2013 8:14:19 PM', '12/9/2013 8:14:46 PM'),
(164, 'admin', '12/9/2013 8:15:51 PM', '12/9/2013 8:16:35 PM'),
(165, 'admin', '12/9/2013 8:20:15 PM', '12/9/2013 8:20:29 PM'),
(166, 'admin', '12/9/2013 8:23:01 PM', '12/9/2013 8:23:25 PM'),
(167, 'admin', '12/9/2013 8:25:54 PM', '12/9/2013 8:26:35 PM'),
(168, 'admin', '12/9/2013 8:26:28 PM', '12/9/2013 8:26:35 PM'),
(169, 'admin', '12/9/2013 8:32:11 PM', '12/9/2013 8:32:26 PM'),
(170, 'admin', '12/9/2013 8:36:42 PM', '12/9/2013 8:37:25 PM'),
(171, 'admin', '12/9/2013 8:37:16 PM', '12/9/2013 8:37:25 PM'),
(172, 'admin', '12/9/2013 8:42:33 PM', '12/9/2013 8:42:45 PM'),
(173, 'admin', '12/9/2013 8:46:52 PM', '12/9/2013 8:47:05 PM'),
(174, 'admin', '12/9/2013 8:50:28 PM', '12/9/2013 8:50:52 PM'),
(175, 'admin', '12/19/2013 3:36:53 PM', '12/19/2013 3:40:40 PM'),
(176, 'admin', '12/19/2013 4:42:53 PM', '12/19/2013 5:42:31 PM'),
(177, 'admin', '12/19/2013 6:22:27 PM', '12/19/2013 6:24:28 PM'),
(178, 'admin', '12/19/2013 6:26:28 PM', '12/19/2013 6:26:41 PM'),
(179, 'admin', '12/19/2013 6:26:51 PM', '12/19/2013 6:27:07 PM'),
(180, 'admin', '12/19/2013 6:27:49 PM', '12/19/2013 6:27:59 PM'),
(181, 'admin', '12/19/2013 6:36:06 PM', '12/19/2013 6:39:11 PM'),
(182, 'admin', '12/19/2013 6:37:10 PM', '12/19/2013 6:39:11 PM'),
(183, 'admin', '12/19/2013 6:39:17 PM', '12/19/2013 6:39:36 PM'),
(184, 'admin', '12/19/2013 6:39:50 PM', '12/19/2013 6:41:27 PM'),
(185, 'admin', '12/19/2013 6:41:03 PM', '12/19/2013 6:41:27 PM'),
(186, 'admin', '12/19/2013 7:04:08 PM', '12/19/2013 7:04:12 PM'),
(187, 'admin', '12/19/2013 8:24:48 PM', '12/19/2013 8:24:51 PM'),
(188, 'admin', '12/19/2013 8:26:02 PM', '12/19/2013 8:26:04 PM'),
(189, 'admin', '12/19/2013 8:26:40 PM', '12/19/2013 8:26:42 PM'),
(190, 'admin', '12/19/2013 8:29:27 PM', '12/19/2013 8:29:33 PM'),
(191, 'admin', '12/19/2013 8:30:47 PM', '12/19/2013 8:30:52 PM'),
(192, 'admin', '12/19/2013 8:31:30 PM', '12/19/2013 8:31:41 PM'),
(193, 'admin', '12/19/2013 8:32:08 PM', '12/19/2013 8:32:44 PM'),
(194, 'admin', '12/19/2013 8:34:05 PM', '12/19/2013 8:34:12 PM'),
(195, 'admin', '12/19/2013 8:34:17 PM', '12/19/2013 8:34:35 PM'),
(196, 'admin', '12/19/2013 8:51:36 PM', '12/19/2013 8:52:00 PM'),
(197, 'admin', '12/19/2013 8:56:36 PM', '12/19/2013 8:56:52 PM'),
(198, 'admin', '12/19/2013 8:57:52 PM', '12/19/2013 8:58:06 PM'),
(199, 'admin', '12/19/2013 9:16:12 PM', '12/19/2013 9:20:54 PM'),
(200, 'admin', '12/19/2013 9:20:58 PM', '12/19/2013 9:21:24 PM'),
(201, 'admin', '12/19/2013 10:10:16 PM', '12/19/2013 10:11:30 PM'),
(202, 'admin', '12/19/2013 10:15:42 PM', '12/19/2013 10:16:50 PM'),
(203, 'admin', '12/19/2013 10:18:24 PM', '12/19/2013 10:18:53 PM'),
(204, 'admin', '12/19/2013 10:19:11 PM', '12/19/2013 10:20:31 PM'),
(205, 'admin', '12/19/2013 10:31:14 PM', '12/19/2013 10:40:40 PM'),
(206, 'admin', '12/19/2013 10:33:01 PM', '12/19/2013 10:40:40 PM'),
(207, 'admin', '12/19/2013 10:35:27 PM', '12/19/2013 10:40:40 PM'),
(208, 'admin', '12/19/2013 10:36:07 PM', '12/19/2013 10:40:40 PM'),
(209, 'admin', '12/19/2013 10:53:11 PM', '12/19/2013 10:54:39 PM'),
(210, 'admin', '12/19/2013 10:55:57 PM', '12/19/2013 10:56:12 PM'),
(211, 'admin', '12/19/2013 10:57:41 PM', '12/19/2013 10:59:02 PM'),
(212, 'admin', '12/19/2013 10:58:55 PM', '12/19/2013 10:59:02 PM'),
(213, 'admin', '12/19/2013 11:18:16 PM', '12/19/2013 11:19:55 PM'),
(214, 'admin', '12/19/2013 11:20:40 PM', '12/19/2013 11:21:00 PM'),
(215, 'admin', '12/19/2013 11:30:38 PM', '12/19/2013 11:31:45 PM'),
(216, 'admin', '12/19/2013 11:34:21 PM', '12/19/2013 11:43:39 PM'),
(217, 'admin', '12/19/2013 11:37:00 PM', '12/19/2013 11:43:39 PM'),
(218, 'admin', '12/19/2013 11:43:19 PM', '12/19/2013 11:43:39 PM'),
(219, 'admin', '12/19/2013 11:48:21 PM', '12/19/2013 11:56:37 PM'),
(220, 'admin', '12/20/2013 12:04:05 AM', '12/20/2013 12:05:45 AM'),
(221, 'admin', '12/20/2013 12:07:53 AM', '12/20/2013 12:22:56 AM'),
(222, 'admin', '12/20/2013 12:22:06 AM', '12/20/2013 12:22:56 AM'),
(223, 'admin', '12/20/2013 12:29:40 AM', '12/20/2013 12:31:29 AM'),
(224, 'admin', '12/20/2013 12:35:12 AM', '12/20/2013 12:37:54 AM'),
(225, 'admin', '12/20/2013 12:39:17 AM', '12/20/2013 12:40:19 AM'),
(226, 'admin', '12/20/2013 12:46:21 AM', '12/20/2013 12:46:36 AM'),
(227, 'admin', '12/20/2013 12:48:16 AM', '12/20/2013 12:48:27 AM'),
(228, 'admin', '12/20/2013 12:49:08 AM', '12/20/2013 12:49:41 AM'),
(229, 'admin', '12/20/2013 12:52:23 AM', '12/20/2013 12:53:25 AM'),
(230, 'admin', '12/20/2013 12:56:49 AM', '12/20/2013 12:58:19 AM'),
(231, 'admin', '12/20/2013 1:09:08 AM', '12/20/2013 1:09:30 AM'),
(232, 'admin', '12/20/2013 1:17:22 AM', '12/20/2013 1:24:38 AM'),
(233, 'admin', '12/20/2013 1:18:11 AM', '12/20/2013 1:24:38 AM'),
(234, 'admin', '12/20/2013 1:19:50 AM', '12/20/2013 1:24:38 AM'),
(235, 'admin', '12/20/2013 1:22:33 AM', '12/20/2013 1:24:38 AM'),
(236, 'admin', '12/20/2013 1:23:24 AM', '12/20/2013 1:24:38 AM'),
(237, 'admin', '12/20/2013 1:25:29 AM', '12/20/2013 1:25:55 AM'),
(238, 'admin', '12/20/2013 1:26:24 AM', '12/20/2013 1:26:38 AM'),
(239, 'admin', '12/20/2013 1:33:03 AM', '12/20/2013 1:33:20 AM'),
(240, 'admin', '12/20/2013 1:34:03 AM', '12/20/2013 1:34:43 AM'),
(241, 'admin', '12/20/2013 1:35:17 AM', '12/20/2013 1:36:05 AM'),
(242, 'admin', '12/20/2013 1:36:28 AM', '12/20/2013 1:37:02 AM'),
(243, 'julie', '12/20/2013 1:37:09 AM', '12/20/2013 1:37:21 AM'),
(244, 'admin', '12/20/2013 1:37:32 AM', '12/20/2013 1:38:00 AM'),
(245, 'adminto', '12/20/2013 1:38:09 AM', '12/20/2013 1:38:12 AM'),
(246, 'adminto', '12/20/2013 1:43:03 AM', '12/20/2013 1:43:51 AM'),
(247, 'admin', '12/20/2013 1:43:25 AM', '12/20/2013 1:44:18 AM'),
(248, 'adminto', '12/20/2013 1:43:59 AM', 'None'),
(249, 'admin', '12/20/2013 1:44:47 AM', '11/3/2010 12:07:59 AM'),
(250, 'admin', '11/3/2010 12:06:27 AM', '11/3/2010 12:07:59 AM'),
(251, 'admin', '11/3/2010 12:09:42 AM', '11/3/2010 12:10:08 AM'),
(252, 'admin', '11/3/2010 12:12:38 AM', '11/3/2010 12:14:20 AM'),
(253, 'admin', '11/3/2010 12:14:49 AM', '12/20/2013 8:19:57 AM'),
(254, 'admin', '12/20/2013 8:16:22 AM', '12/20/2013 8:19:57 AM'),
(255, 'admin', '12/20/2013 8:20:52 AM', '12/20/2013 8:21:06 AM'),
(256, 'admin', '12/20/2013 8:24:19 AM', '12/20/2013 8:26:45 AM'),
(257, 'admin', '12/20/2013 8:26:50 AM', '12/20/2013 8:28:13 AM'),
(258, 'admin', '12/20/2013 8:31:30 AM', '12/20/2013 8:33:05 AM'),
(259, 'admin', '12/20/2013 8:31:54 AM', '12/20/2013 8:33:05 AM'),
(260, 'admin', '12/20/2013 8:32:42 AM', '12/20/2013 8:33:05 AM'),
(261, 'admin', '12/20/2013 8:36:13 AM', '12/20/2013 8:39:07 AM'),
(262, 'admin', '12/20/2013 8:42:15 AM', '12/20/2013 8:43:06 AM'),
(263, 'admin', '12/20/2013 8:47:36 AM', '12/20/2013 8:48:21 AM'),
(264, 'admin', '12/20/2013 8:50:53 AM', '12/20/2013 8:51:55 AM'),
(265, 'admin', '12/20/2013 8:54:25 AM', '12/20/2013 8:56:59 AM'),
(266, 'admin', '12/20/2013 9:03:29 AM', '12/20/2013 9:04:25 AM'),
(267, 'admin', '12/20/2013 9:05:06 AM', '12/20/2013 9:05:31 AM'),
(268, 'admin', '12/20/2013 9:06:24 AM', '12/20/2013 9:06:43 AM'),
(269, 'admin', '12/20/2013 9:07:53 AM', '12/20/2013 9:08:27 AM'),
(270, 'admin', '12/20/2013 9:09:23 AM', '12/20/2013 9:10:06 AM'),
(271, 'admin', '12/20/2013 9:10:33 AM', '12/20/2013 9:12:03 AM'),
(272, 'admin', '12/20/2013 9:14:07 AM', '12/20/2013 9:21:26 AM'),
(273, 'admin', '12/20/2013 9:40:38 AM', '12/20/2013 9:40:48 AM'),
(274, 'admin', '12/20/2013 9:41:29 AM', '12/20/2013 9:43:53 AM'),
(275, 'admin', '12/20/2013 9:42:24 AM', '12/20/2013 9:43:53 AM'),
(276, 'admin', '12/20/2013 9:43:29 AM', '12/20/2013 9:43:53 AM'),
(277, 'julie', '12/20/2013 9:44:39 AM', '12/20/2013 9:44:46 AM'),
(278, 'admin', '12/20/2013 9:46:26 AM', '12/20/2013 9:46:29 AM'),
(279, 'admin', '12/20/2013 9:47:02 AM', '12/20/2013 9:47:19 AM'),
(280, 'admin', '12/20/2013 9:47:40 AM', '12/20/2013 9:47:42 AM'),
(281, 'admin', '12/20/2013 9:51:37 AM', '12/20/2013 9:52:09 AM'),
(282, 'admin', '12/20/2013 10:13:42 AM', '12/20/2013 10:16:47 AM'),
(283, 'admin', '12/20/2013 10:18:02 AM', '12/20/2013 10:18:11 AM'),
(284, 'admin', '12/20/2013 10:19:47 AM', '12/20/2013 10:21:40 AM'),
(285, 'admin', '12/20/2013 10:21:32 AM', '12/20/2013 10:21:40 AM'),
(286, 'admin', '12/20/2013 10:24:19 AM', '12/20/2013 10:24:56 AM'),
(287, 'admin', '12/20/2013 10:25:29 AM', '12/20/2013 10:25:37 AM');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_order`
--

CREATE TABLE IF NOT EXISTS `tbl_order` (
  `Order_ID` int(100) NOT NULL AUTO_INCREMENT,
  `Order_Date` varchar(100) NOT NULL,
  `Supplier_Name` varchar(100) NOT NULL,
  `Person_In_Charge` varchar(100) NOT NULL,
  `Product_ID` varchar(100) NOT NULL,
  `Quantity` varchar(100) NOT NULL,
  `Total` varchar(100) NOT NULL,
  `Remark` varchar(100) NOT NULL,
  `Expected_Delivery` varchar(100) NOT NULL,
  PRIMARY KEY (`Order_ID`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=3 ;

--
-- Dumping data for table `tbl_order`
--

INSERT INTO `tbl_order` (`Order_ID`, `Order_Date`, `Supplier_Name`, `Person_In_Charge`, `Product_ID`, `Quantity`, `Total`, `Remark`, `Expected_Delivery`) VALUES
(1, '12/6/2013 1:28:54 AM', 'Miguel Carlo Pascua', 'Miguel Carlo Pascua', 'P-000001', '10', ' 50000', 'Accepted', '12/16/2013 1:28:54 AM'),
(2, '12/6/2013 2:13:58 AM', 'Batangas Furniture Shop', 'Juan Dela Cruz', 'P-000001', '5', '30000', 'Pending', '12/6/2013 2:13:58 AM');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_password`
--

CREATE TABLE IF NOT EXISTS `tbl_password` (
  `ID` int(100) NOT NULL AUTO_INCREMENT,
  `Password` varchar(100) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=2 ;

--
-- Dumping data for table `tbl_password`
--

INSERT INTO `tbl_password` (`ID`, `Password`) VALUES
(1, 'furniture');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_product`
--

CREATE TABLE IF NOT EXISTS `tbl_product` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Product_ID` varchar(200) NOT NULL,
  `Product_Name` varchar(200) NOT NULL,
  `Category` varchar(200) NOT NULL,
  `Brand` varchar(200) NOT NULL,
  `Description` varchar(200) NOT NULL,
  `Initial_Supplier` varchar(200) NOT NULL,
  `Cost` varchar(200) NOT NULL,
  `Quantity` int(200) NOT NULL,
  `Unit_Price` varchar(200) NOT NULL,
  `Critical_Point` int(200) NOT NULL,
  `Remark` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;

--
-- Dumping data for table `tbl_product`
--

INSERT INTO `tbl_product` (`No`, `Product_ID`, `Product_Name`, `Category`, `Brand`, `Description`, `Initial_Supplier`, `Cost`, `Quantity`, `Unit_Price`, `Critical_Point`, `Remark`) VALUES
(1, 'P-000001', 'White Sofa', 'Sofa', 'Uratex', 'White sofa', 'Batangas Furniture Shop', '5000', 5, '10000', 3, 'Active'),
(2, 'P-000002', 'Wooden Sofa', 'Sofa', 'A', 'wooden sofa', 'Batangas Furniture Shop', '10000', 3, '15000', 1, 'Pull-Out'),
(3, 'P-000003', 'Brown Dining Table', 'Dining Table', 'Sample', 'brown dining table', 'Batangas Furniture Shop', '10000', 2, '15000', 3, 'Active');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_purchase`
--

CREATE TABLE IF NOT EXISTS `tbl_purchase` (
  `Purchase_ID` int(100) NOT NULL AUTO_INCREMENT,
  `Purchase_Date` varchar(100) NOT NULL,
  `Customer_Name` varchar(100) NOT NULL,
  `Person_In_Charge` varchar(100) NOT NULL,
  `Product_ID` varchar(100) NOT NULL,
  `Quantity` varchar(100) NOT NULL,
  `Total` varchar(100) NOT NULL,
  `Remark` varchar(100) NOT NULL,
  `Expected_Delivery` varchar(100) NOT NULL,
  PRIMARY KEY (`Purchase_ID`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=9 ;

--
-- Dumping data for table `tbl_purchase`
--

INSERT INTO `tbl_purchase` (`Purchase_ID`, `Purchase_Date`, `Customer_Name`, `Person_In_Charge`, `Product_ID`, `Quantity`, `Total`, `Remark`, `Expected_Delivery`) VALUES
(1, '12/9/2013 4:14:56 PM', 'Home Plus', 'Pedro Magtibay', 'P-000001', '10', ' 100000', 'Delivered', '12/9/2013 4:14:56 PM'),
(2, '12/19/2013 10:31:17 PM', 'Home Plus', 'Pedro Magtibay', 'P-000001', '1', ' 10000', 'Pick-up', '12/19/2013 10:31:17 PM'),
(3, '12/19/2013 10:33:04 PM', 'Home Plus', 'Pedro Magtibay', 'P-000001', '1', ' 10000', 'Delivered', '12/19/2013 10:33:04 PM'),
(4, '12/19/2013 10:36:09 PM', 'Home Plus', 'Pedro Magtibay', 'P-000001', '1', ' 10000', 'Delivered', '12/19/2013 10:36:09 PM'),
(6, '12/19/2013 10:53:59 PM', 'Perlin Diloy', 'Perlin Diloy', 'P-000001', '1', ' 10000', 'Pick-up', '12/19/2013 10:53:59 PM'),
(7, '12/20/2013 12:04:13 AM', 'Perlin Diloy', 'Perlin Diloy', 'P-000001', '1', ' 10000', 'Delivered', '12/20/2013 12:04:13 AM'),
(8, '12/20/2013 8:54:27 AM', 'Home Plus', 'Pedro Magtibay', 'P-000003', '1', ' 15000', 'Delivered', '12/20/2013 8:54:27 AM');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_supplier`
--

CREATE TABLE IF NOT EXISTS `tbl_supplier` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Supplier_ID` varchar(200) NOT NULL,
  `Supplier_Name` varchar(200) NOT NULL,
  `Representative1` varchar(200) NOT NULL,
  `Representative2` varchar(200) NOT NULL,
  `Mobile_Number` varchar(200) NOT NULL,
  `Address` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=3 ;

--
-- Dumping data for table `tbl_supplier`
--

INSERT INTO `tbl_supplier` (`No`, `Supplier_ID`, `Supplier_Name`, `Representative1`, `Representative2`, `Mobile_Number`, `Address`) VALUES
(1, 'S-00001', 'Batangas Furniture Shop', 'Juan Dela Cruz', '', '091111111', 'Batangas City, Cavite'),
(2, 'S-00002', 'Miguel Carlo Pascua', 'Miguel Carlo Pascua', '', '09111111', 'Cavite City');

-- --------------------------------------------------------

--
-- Table structure for table `tbl_temp_order`
--

CREATE TABLE IF NOT EXISTS `tbl_temp_order` (
  `Order_ID` varchar(100) NOT NULL,
  `Order_Date` varchar(100) NOT NULL,
  `Supplier_Name` varchar(100) NOT NULL,
  `Person_In_Charge` varchar(100) NOT NULL,
  `Product_ID` varchar(100) NOT NULL,
  `Quantity` varchar(100) NOT NULL,
  `Total` varchar(100) NOT NULL,
  `Remark` varchar(100) NOT NULL,
  `Expected_Delivery` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tbl_temp_purchase`
--

CREATE TABLE IF NOT EXISTS `tbl_temp_purchase` (
  `Purchase_ID` varchar(100) NOT NULL,
  `Purchase_Date` varchar(100) NOT NULL,
  `Customer_Name` varchar(100) NOT NULL,
  `Person_In_Charge` varchar(100) NOT NULL,
  `Product_ID` varchar(100) NOT NULL,
  `Quantity` varchar(100) NOT NULL,
  `Total` varchar(100) NOT NULL,
  `Remark` varchar(100) NOT NULL,
  `Expected_Delivery` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `tbl_users`
--

CREATE TABLE IF NOT EXISTS `tbl_users` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `ID` varchar(100) NOT NULL,
  `Lastname` varchar(100) NOT NULL,
  `Firstname` varchar(100) NOT NULL,
  `Middlename` varchar(100) NOT NULL,
  `MobileNumber` varchar(100) NOT NULL,
  `Address` varchar(100) NOT NULL,
  `Username` varchar(100) NOT NULL,
  `Usertype` varchar(100) NOT NULL,
  `Password` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB  DEFAULT CHARSET=latin1 AUTO_INCREMENT=3 ;

--
-- Dumping data for table `tbl_users`
--

INSERT INTO `tbl_users` (`No`, `ID`, `Lastname`, `Firstname`, `Middlename`, `MobileNumber`, `Address`, `Username`, `Usertype`, `Password`) VALUES
(1, 'Em-0001', 'Admin', 'Admin', 'Admin', '00001', 'Cavite', 'admin', 'Administrator', 'admin'),
(2, 'Em-0002', 'Tamargo', 'Julie Ann', 'A', '11', 'Cavite City', 'julie', 'Assistant Administrator', 'julie');

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
