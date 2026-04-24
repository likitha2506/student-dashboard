async function renderCharts() {
    const res = await fetch('/api/charts');
    const data = await res.json();

    new Chart(document.getElementById('marksChart'), {
        type: 'bar',
        data: {
            labels: data.marksBySubject.map(x => x.subject),
            datasets: [{ label: 'Average Marks', data: data.marksBySubject.map(x => x.avg_marks), backgroundColor: '#0d6efd' }]
        }
    });

    new Chart(document.getElementById('attendanceChart'), {
        type: 'doughnut',
        data: {
            labels: data.attendancePercent.map(x => x.name),
            datasets: [{ label: 'Attendance %', data: data.attendancePercent.map(x => x.attendance_percent), backgroundColor: ['#198754', '#0dcaf0', '#ffc107', '#dc3545', '#6f42c1', '#fd7e14'] }]
        }
    });

    new Chart(document.getElementById('gpaChart'), {
        type: 'line',
        data: {
            labels: data.gpaTrend.map(x => x.month),
            datasets: [{ label: 'GPA Trend', data: data.gpaTrend.map(x => x.gpa), borderColor: '#6610f2', fill: false, tension: 0.3 }]
        }
    });

    new Chart(document.getElementById('deptChart'), {
        type: 'bar',
        data: {
            labels: data.departmentPerformance.map(x => x.department),
            datasets: [{ label: 'Department Avg Marks', data: data.departmentPerformance.map(x => x.avg_marks), backgroundColor: '#20c997' }]
        }
    });
}

renderCharts();
