CREATE TABLE tbl_GoogleMapsLeads ( LeadID AUTOINCREMENT PRIMARY KEY, PlaceName TEXT(255), FullAddress MEMO, PhoneNumber TEXT(50), Website TEXT(255), Rating DOUBLE, ApifyRunID TEXT(50), DateImported DATETIME DEFAULT Now() );

CREATE TABLE tbl_APICredits ( LogID AUTOINCREMENT PRIMARY KEY, ActorName TEXT(100), CreditsUsed DOUBLE, RemainingBalance DOUBLE, LogDate DATETIME DEFAULT Now() );