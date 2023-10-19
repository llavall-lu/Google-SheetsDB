<?php
include 'db_connect.php';

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $userID = $_POST['UserID'];
    $firstName = $_POST['FirstName'];
    $surname = $_POST['Surname'];
    $address = $_POST['Address'];
    $email = $_POST['Email'];
    $phoneNumber = $_POST['PhoneNumber'];
    $dateOfBirth = $_POST['DateOfBirth'];
    $country = $_POST['CountryOfResidence'];
    $nationality = $_POST['Nationality'];
    $status = $_POST['EmploymentStatus'];
    $department = $_POST['Department'];
    $role = $_POST['Role'];
    $startDate = $_POST['StartDate'];
    $endDate = $_POST['EndDate'];

    $sql = "UPDATE PersonalInformation SET FirstName=?, Surname=?, Address=?, Email=?, PhoneNumber=?, DateOfBirth=?, CountryOfResidence=?, Nationality=?, EmploymentStatus=?, Department=?, Role=?, StartDate=?, EndDate=? WHERE UserID=?";
    $stmt = $conn->prepare($sql);
    $stmt->bind_param("sssssssssssssi", $firstName, $surname, $address, $email, $phoneNumber, $dateOfBirth, $country, $nationality, $status, $department, $role, $startDate, $endDate, $userID);
    $stmt->execute();

    header("Location: index.php?updated=true"); // Redirect to main page with the query parameter
    exit; // Stop further processing
}

if (!isset($_GET['UserID'])) {
    die("UserID not provided.");
}
$userID = $_GET['UserID'];
$sql = "SELECT * FROM PersonalInformation WHERE UserID=?";
$stmt = $conn->prepare($sql);
$stmt->bind_param("i", $userID);
$stmt->execute();
$result = $stmt->get_result();
$userDetails = $result->fetch_assoc();
?>

<!DOCTYPE html>
<html>
<head>
    <title>Edit User Details</title>
    <link rel="stylesheet" type="text/css" href="styles.css">
</head>
<body>
    <h2>Edit User Details</h2>
    <form action="edit.php" method="post">
    <input type="hidden" name="UserID" value="<?php echo $userDetails['UserID']; ?>">
    First Name: <input type="text" name="FirstName" value="<?php echo $userDetails['FirstName']; ?>"><br>
    Surname: <input type="text" name="Surname" value="<?php echo $userDetails['Surname']; ?>"><br>
    Address: <textarea name="Address"><?php echo $userDetails['Address']; ?></textarea><br>
    Email: <input type="email" name="Email" value="<?php echo $userDetails['Email']; ?>"><br>
    Phone Number: <input type="text" name="PhoneNumber" value="<?php echo $userDetails['PhoneNumber']; ?>"><br>
    Date of Birth: <input type="date" name="DateOfBirth" value="<?php echo $userDetails['DateOfBirth']; ?>"><br>
    Country: <input type="text" name="CountryOfResidence" value="<?php echo $userDetails['CountryOfResidence']; ?>"><br>
    Nationality: <input type="text" name="Nationality" value="<?php echo $userDetails['Nationality']; ?>"><br>
    Employment Status: <input type="text" name="EmploymentStatus" value="<?php echo $userDetails['EmploymentStatus']; ?>"><br>
    Department: <input type="text" name="Department" value="<?php echo $userDetails['Department']; ?>"><br>
    Role: <input type="text" name="Role" value="<?php echo $userDetails['Role']; ?>"><br>
    Start Date: <input type="date" name="StartDate" value="<?php echo $userDetails['StartDate']; ?>"><br>
    End Date: <input type="date" name="EndDate" value="<?php echo $userDetails['EndDate']; ?>"><br>
    <input type="submit" value="Update">
</form>

    <br>
    <a href="index.php">Back to list</a>
    
</body>
</html>
