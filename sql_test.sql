UPDATE Persons
SET ContractAddress = CONCAT(Address.Address1, ' ', Address.Address2, ' ', Address.Address3)
FROM Address
WHERE Persons.PersonID = Address.PersonID;