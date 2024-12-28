---------- CREATE DATABASE ----------

-- Create database called 'Visionboarddb'
-- Connected to 'master' db
USE master
GO
-- Create the new database if it does not exist already
IF NOT EXISTS (
    SELECT [name]
        FROM sys.databases
        WHERE [name] = N'Visionboarddb'
)
CREATE DATABASE Visionboarddb
GO

USE Visionboarddb
---------------------------------------


---------- GOALS ----------
-- CREATE [Goals] table in schema [dbo]
-- Drop the table if it already exists
IF OBJECT_ID('[dbo].[Goals]', 'U') IS NOT NULL
DROP TABLE [dbo].[Goals]
GO
-- Create the table in the specified schema
CREATE TABLE [dbo].[Goals]
(
    [IdGoal] INT NOT NULL IDENTITY(1,1), --identity (inicia en 1, y aumente de 1en1)
    [title] VARCHAR(250) NOT NULL UNIQUE,
	[description] TEXT NOT NULL,
    [completionMonth] TINYINT NOT NULL, -- serian los meses del año
    [imageUrl] VARCHAR(255) NOT NULL,

    --Constraints
    CONSTRAINT PK_Goals PRIMARY KEY 
    (
        IdGoal ASC
    ),
);
GO

----------- INSERTION ---------------
--INSERT INTO Goals (title, description, completionMonth, imageUrl)
--VALUES (
	--'dominar vb6',
	--'es la unica meta', 
	--5, 
	--'C:\Users\yourimageroutehere.jpg');

SELECT * FROM Goals;

