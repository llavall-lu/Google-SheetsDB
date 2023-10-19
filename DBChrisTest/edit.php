<?php
include 'db_connect.php';

$isModal = isset($_GET['modal']) && $_GET['modal'] == 'true';

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

    header('Content-Type: application/json');
    echo json_encode(['updated' => true]);
    exit();
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

function renderFormFields($userDetails) {
    // This function will render the form fields for both modal and full-page views.
?>
    <input type="hidden" name="UserID" value="<?php echo htmlspecialchars($userDetails['UserID']); ?>">
    
    <div class="form-row">
        <label for="FirstName">First Name:</label>
        <input type="text" id="FirstName" name="FirstName" value="<?php echo $userDetails['FirstName']; ?>">
    </div>

    <div class="form-row">
        <label for="Surname">Surname:</label>
        <input type="text" id="Surname" name="Surname" value="<?php echo $userDetails['Surname']; ?>">
    </div>

    <div class="form-row">
        <label for="Address">Address:</label>
        <textarea id="Address" name="Address"><?php echo $userDetails['Address']; ?></textarea>
    </div>

    <div class="form-row">
        <label for="Email">Email:</label>
        <input type="email" id="Email" name="Email" value="<?php echo $userDetails['Email']; ?>">
    </div>

    <div class="form-row">
        <label for="PhoneNumber">Phone Number:</label>
        <input type="text" id="PhoneNumber" name="PhoneNumber" value="<?php echo $userDetails['PhoneNumber']; ?>">
    </div>

    <div class="form-row">
        <label for="DateOfBirth">Date of Birth:</label>
        <input type="date" id="DateOfBirth" name="DateOfBirth" value="<?php echo $userDetails['DateOfBirth']; ?>">
    </div>

    <div class="form-row">
        <label for="CountryOfResidence">Country:</label>
        <input type="text" id="CountryOfResidence" name="CountryOfResidence" value="<?php echo $userDetails['CountryOfResidence']; ?>">
    </div>

    <div class="form-row">
        <label for="Nationality">Nationality:</label>
        <input type="text" id="Nationality" name="Nationality" value="<?php echo $userDetails['Nationality']; ?>">
    </div>

    <div class="form-row">
        <label for="EmploymentStatus">Employment Status:</label>
        <input type="text" id="EmploymentStatus" name="EmploymentStatus" value="<?php echo $userDetails['EmploymentStatus']; ?>">
    </div>

    <div class="form-row">
        <label for="Department">Department:</label>
        <input type="text" id="Department" name="Department" value="<?php echo $userDetails['Department']; ?>">
    </div>

    <div class="form-row">
        <label for="Role">Role:</label>
        <input type="text" id="Role" name="Role" value="<?php echo $userDetails['Role']; ?>">
    </div>

    <div class="form-row">
        <label for="StartDate">Start Date:</label>
        <input type="date" id="StartDate" name="StartDate" value="<?php echo $userDetails['StartDate']; ?>">
    </div>

    <div class="form-row">
        <label for="EndDate">End Date:</label>
        <input type="date" id="EndDate" name="EndDate" value="<?php echo $userDetails['EndDate']; ?>">
    </div>

    <div class="form-row">
        <input type="submit" value="Update">
    </div>

<?php
}

if ($isModal) {
    renderFormFields($userDetails);
    exit();
}

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
        <?php renderFormFields($userDetails); ?>
    </form>

    <br>
    <a href="index.php">Back to list</a>
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            const form = document.querySelector("form");

            form.addEventListener('submit', async function(e) {
                e.preventDefault();
                
                // Assuming you will fetch data here, but it's not in the code you provided
                try {
                    let response = await fetch(/* fetch URL and options */);
                    let data = await response.json();
                    
                    if (data.updated === true) {
                        alert("Data updated successfully!");
                        
                        if (window.parent && window.parent.$) {
                            window.parent.$('#editModal').hide();
                            window.parent.$("body").removeClass("no-scroll");
                        }

                        window.parent.location.reload();
                    } else {
                        alert("Error updating data. Please try again.");
                    }
                } catch (error) {
                    console.error("Error updating data", error);
                }
            });
        });
    </script>

</body>
</html>
