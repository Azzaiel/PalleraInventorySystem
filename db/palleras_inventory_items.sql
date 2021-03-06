CREATE DATABASE  IF NOT EXISTS `palleras_inventory` /*!40100 DEFAULT CHARACTER SET latin1 */;
USE `palleras_inventory`;
-- MySQL dump 10.13  Distrib 5.6.13, for Win32 (x86)
--
-- Host: 127.0.0.1    Database: palleras_inventory
-- ------------------------------------------------------
-- Server version	5.6.12-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `items`
--

DROP TABLE IF EXISTS `items`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `items` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `Supplier_ID` int(11) DEFAULT NULL,
  `Item_type_ID` int(11) DEFAULT NULL,
  `Name` varchar(255) DEFAULT NULL,
  `Retail_Price` int(11) DEFAULT NULL,
  `Unit_Price` int(11) DEFAULT NULL,
  `Item_Code` varchar(100) DEFAULT NULL,
  `Created_By` varchar(255) DEFAULT NULL,
  `Created_Date` datetime DEFAULT NULL,
  `Last_Mod_By` varchar(255) DEFAULT NULL,
  `Last_Mod_Date` datetime DEFAULT NULL,
  `ACTIVE` varchar(1) DEFAULT NULL,
  `bar_code` varchar(400) DEFAULT NULL,
  `quantity` int(11) DEFAULT NULL,
  `CRITICAL_LEVEL` int(11) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `ID_UNIQUE` (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `items`
--

LOCK TABLES `items` WRITE;
/*!40000 ALTER TABLE `items` DISABLE KEYS */;
INSERT INTO `items` VALUES (1,2,2,'Flat White 122 Liter 1',100,131,'DV 500','System','2014-01-25 21:38:24','icha','2014-03-08 12:04:03','Y',NULL,34,4),(2,2,1,'Black 1233',120,150,'DV 501','System','2014-02-22 15:07:19',NULL,'2014-03-08 12:01:05','Y',NULL,2,5),(4,3,4,'Plywood',200,250,'MC 101','System','2014-02-23 10:50:15',NULL,'2014-03-08 12:04:21','Y',NULL,45,3),(5,3,5,'Pink Flower Tile',120,170,'MC 100','System','2014-02-23 10:51:20',NULL,'2014-03-08 12:04:26','Y',NULL,19,4),(6,0,3,'Model 123',2000,2800,'MM 100','System','2014-02-23 10:51:57',NULL,'2014-02-23 13:56:00','Y',NULL,NULL,NULL),(7,0,6,'Led Yellow',80,120,'MM 101','System','2014-02-23 10:53:32',NULL,'2014-02-23 13:56:03','Y',NULL,NULL,NULL),(8,2,1,'Pink Flat',500,555,'DV 502','System','2014-02-23 10:55:03',NULL,'2014-03-08 12:03:11','Y',NULL,NULL,2),(9,1,3,'Model 200',12000,15000,'MM 102','System','2014-02-23 10:56:08',NULL,'2014-03-08 12:01:15','Y',NULL,10,5),(10,2,2,'Brown',200,230,'DV 503','System','2014-02-23 10:57:57',NULL,'2014-03-08 12:04:09','Y',NULL,NULL,3),(11,2,7,'Metalic Red',500,620,'DC 504','System','2014-02-23 11:02:39',NULL,'2014-03-08 12:04:15','Y',NULL,NULL,5),(12,1,6,'Disco Lights',200,300,'MM 300','System','2014-03-08 12:06:20',NULL,'2014-03-08 12:06:20','Y',NULL,0,2);
/*!40000 ALTER TABLE `items` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-03-09 17:20:10
