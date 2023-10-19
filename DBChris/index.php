<?php
// Include the database connection
include 'db_connect.php';

// Check if the 'updated' query parameter is set and true
$showUpdateMessage = isset($_GET['updated']) && $_GET['updated'] == 'true';
?>

<!DOCTYPE html>
<html>
<head>
    <title>Employee Details</title>
    <link rel="stylesheet" type="text/css" href="styles.css">
</head>
<body>
    <h2>Employee Details</h2>
    <table>
    <thead>
    <tr>
        <th>ID</th>
        <th>First Name</th>
        <th>Surname</th>
        <th>Address</th>
        <th>Email</th>
        <th>Phone Number</th>
        <th>Date of Birth</th>
        <th>Country</th>
        <th>Nationality</th>
        <th>Status</th>
        <th>Department</th>
        <th>Role</th>
        <th>Start Date</th>
        <th>End Date</th>
        <th>Actions</th>
    </tr>
</thead>
<tbody>
    <?php
    $sql = "SELECT * FROM PersonalInformation";
    $result = $conn->query($sql);

    if ($result->num_rows > 0) {
        while($row = $result->fetch_assoc()) {
            echo "<tr>";
            echo "<td>" . $row["UserID"] . "</td>";
            echo "<td>" . $row["FirstName"] . "</td>";
            echo "<td>" . $row["Surname"] . "</td>";
            echo "<td>" . $row["Address"] . "</td>";
            echo "<td>" . $row["Email"] . "</td>";
            echo "<td>" . $row["PhoneNumber"] . "</td>";
            echo "<td>" . $row["DateOfBirth"] . "</td>";
            echo "<td>" . $row["CountryOfResidence"] . "</td>";
            echo "<td>" . $row["Nationality"] . "</td>";
            echo "<td>" . $row["EmploymentStatus"] . "</td>";
            echo "<td>" . $row["Department"] . "</td>";
            echo "<td>" . $row["Role"] . "</td>";
            echo "<td>" . $row["StartDate"] . "</td>";
            echo "<td>" . $row["EndDate"] . "</td>";
            echo "<td><a href='edit.php?UserID=" . $row["UserID"] . "'>Edit</a></td>";
            echo "</tr>";
        }
    } else {
        echo "<tr><td colspan='15'>No records found</td></tr>";
    }
    $conn->close();
    ?>
</tbody>
    </table>
    <?php if ($showUpdateMessage): ?>
    <script>
        alert("Details updated successfully!");
        // Remove the query parameter from the URL without reloading the page
        if (window.history.replaceState) {
            window.history.replaceState({}, document.title, window.location.pathname);
        }
    </script>
<?php endif; ?>

</body>
</html>
