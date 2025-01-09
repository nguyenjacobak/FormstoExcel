function updateProjects() {
    const name = document.getElementById('name').value;
    const projectType = document.getElementById('project_type').value;
    if (name && projectType) {
        window.location.href = `?name=${name}&project_type=${projectType}`;
    }
}

function goToNextForm() {
  const name = document.getElementById('name').value;
  const projectType = document.getElementById('project_type').value;
  const projectName = document.getElementById('project_list').value;
  const formType = document.querySelector('input[name="form_type"]:checked').value;
  console.log(name, projectType, projectName, formType);

  if (name && projectType && projectName && formType) {
      let data = {
            name: name,
            projectType: projectType,
            projectName: projectName,
            formType: formType
      }

      let nextPage = '';
      switch (formType) {
          case 'hoiDongChuyenMon':
              nextPage = 'hoiDongChuyenMon';
              break;
          case 'canBoHuongDan1':
              nextPage = 'baoCaoTienDoL1';
              break;
          case 'canBoHuongDan2':
              nextPage = 'baoCaoTienDoL2';
              break;
          case 'canBoHuongDan3':
              nextPage = 'huongdan3';
              break;
          case 'canBoPhanBien':
              nextPage = 'canBoPhanBien';
              break;
      }

      fetch(`/get-students?project_name=${projectName}`)
              .then(response => response.json())
              .then(students => {
                console.log(students);
                  data['students'] = students;
                  localStorage.setItem('data', JSON.stringify(data));
                  window.location.href = `${nextPage}`;
              });
  } else {
      alert('Please fill in all required fields.');
  }
}