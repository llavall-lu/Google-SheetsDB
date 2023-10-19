<?php
// Include the database connection
include 'db_connect.php';
?>

<!DOCTYPE html>
<html>
<head>
    <title>Employee Details</title>
    <link rel="stylesheet" type="text/css" href="styles.css">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
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
<tbody id="employeeTableBody"> <!-- Assign an ID for easy reference -->
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
 <!-- Edit Modal -->
 <div id="editModal" class="modal">
        <div class="modal-content">
            <span class="close">&times;</span>
            <h2>Edit User Details</h2>
            <form id="editForm" action="edit.php" method="post">
                <!-- Form content will be loaded dynamically -->
            </form>
        </div>
    </div>


    <script>
    $(document).ready(function(){
        var modal = $('#editModal');
        var closeModal = $('.close');

        // Function to reload the employee table data
        function loadEmployeeData() {
            $.get('index.php', function(data){
                // Parse the fetched data for the table body content
                let fetchedBody = $(data).find('#employeeTableBody').html();
                $('#employeeTableBody').html(fetchedBody);
            });
        }

        $(document).on('click', 'a[href^="edit.php"]', function(e){
            e.preventDefault();
            var userID = $(this).attr('href').split('=')[1];
            $.get('edit.php', {UserID: userID, modal: true}, function(data){
                $('#editForm').html(data);
                modal.show();
                $("body").addClass("no-scroll");  
            });
        });

        closeModal.click(function(){
            modal.hide();
            $("body").removeClass("no-scroll");
        });

        $(document).on('submit', '#editForm', function(e){
            e.preventDefault();
            $.post($(this).attr('action'), $(this).serialize(), function(data){
                modal.hide();
                $("body").removeClass("no-scroll");
                if(data.updated){
                    alert("Details updated successfully!");
                    loadEmployeeData();
                } else {
                    alert("Error updating details. Please try again.");
                }
            }, 'json');
        });
    });
    </script>
</body>
</html>
