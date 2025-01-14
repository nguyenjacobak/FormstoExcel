export function getCookie(name) {
    const cookies = document.cookie.split(';');
    for (let cookie of cookies) {
        const [key, value] = cookie.trim().split('=');
        if (key === name) return decodeURIComponent(value);
    }
    return null;
}

export function submitForm(url){
    document.getElementById('evaluationForm').addEventListener('submit', async function(e) {
        e.preventDefault(); 

        const form = e.target;
        const formData = new FormData(form);

        try {
            const response = await fetch(`${url}`, {
                method: "POST",
                headers: {
                    "X-CSRFToken": formData.get('csrfmiddlewaretoken') 
                },
                body: formData
            });

            if (response.ok) {
                const data = await response.json();
                alert(data.message); 
                window.location.href = "/"; 
            } else {
                const error = await response.json();
                alert(error.error || "Đã có lỗi xảy ra, vui lòng thử lại!");
            }
        } catch (error) {
            console.error("Error:", error);
            alert("Đã có lỗi xảy ra, vui lòng thử lại!");
        }
    });
}

export function autoFillGPA(){
    const inputs = document.querySelectorAll(".diem");
    const weights = document.querySelectorAll(".weight");
    inputs.forEach((input, index) => {
        input.addEventListener("input", function () {
            const studentId = parseInt(input.getAttribute('stt').slice(2)); // Lấy ID sinh viên
            const studentInputs = document.querySelectorAll(`.diemSV${studentId}`); // Lấy tất cả ô điểm của sinh viên
            const gpaInput = document.getElementById(`diemGPASV${studentId}`);
            let total = 0;
            let count = 0;

            studentInputs.forEach((i, index) => {
                const value = parseFloat(i.value);
                if (!isNaN(value)) {
                    const weight = parseFloat(weights[0].innerHTML.replace("%", "")) / 100
                    total += value * weight;
                    count++;
                }
            });
            if (count === studentInputs.length) {
                gpaInput.value = (total).toFixed(1); // Làm tròn đến 2 chữ số thập phân
            } else {
                gpaInput.value = ""; // Nếu chưa đủ 4 ô, xóa giá trị GPA
            }
        });
    });
}