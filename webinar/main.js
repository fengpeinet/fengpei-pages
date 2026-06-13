function smoothScrollTo(target, duration = 800) {
  const start = window.scrollY;
  const end = target.getBoundingClientRect().top + window.scrollY;
  const distance = end - start;
  let startTime = null;

  function easeInOutQuad(t) {
    return t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t;
  }

  function step(timestamp) {
    if (!startTime) startTime = timestamp;
    const elapsed = timestamp - startTime;
    const progress = Math.min(elapsed / duration, 1);
    window.scrollTo(0, start + distance * easeInOutQuad(progress));
    if (progress < 1) requestAnimationFrame(step);
  }

  requestAnimationFrame(step);
}

document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.scroll-cta, .form-float-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.preventDefault();
      const target = document.getElementById('form');
      if (target) smoothScrollTo(target, 600);
    });
  });
  if (new URLSearchParams(window.location.search).has('preview')) {
    document.getElementById('registrationForm').style.display = 'none';
    document.getElementById('successMessage').style.display = 'block';
  }
  const form = document.getElementById('registrationForm');
  const submitBtn = document.getElementById('submitBtn');
  const btnText = submitBtn.querySelector('.btn-text');
  const btnLoader = submitBtn.querySelector('.btn-loader');
  const successMessage = document.getElementById('successMessage');

  // 您的 Google Apps Script Web App 網址
  const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyZFb_XfG8rvCmVUCD-fujdxCEn2S9AUNP7qtqgNVwkWPI0QhtuFWsZ19C2zP8W60pb/exec';

  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    // 取得資料
    const formData = new FormData(form);
    const data = Object.fromEntries(formData.entries());

    // UI 改變為 Loading 狀態
    submitBtn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline';

    try {
      // 傳送資料到 GAS (Google Apps Script)
      await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
      });

      // 假定請求成功
      form.style.display = 'none';
      successMessage.style.display = 'block';

      // 平滑滾動到成功訊息位置
      successMessage.scrollIntoView({ behavior: 'smooth', block: 'center' });

    } catch (error) {
      console.error('Error submitting form:', error);
      alert('報名發生錯誤，請稍後再試或是聯絡承辦單位！');

      // 恢復按鈕狀態
      submitBtn.disabled = false;
      btnText.style.display = 'inline';
      btnLoader.style.display = 'none';
    }
  });

  // 為其他人報名功能
  const resetBtn = document.getElementById('resetBtn');
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      form.reset(); // 清空表單
      successMessage.style.display = 'none'; // 隱藏成功訊息
      form.style.display = 'block'; // 顯示表單
      
      // 恢復送出按鈕狀態
      submitBtn.disabled = false;
      btnText.style.display = 'inline';
      btnLoader.style.display = 'none';
      
      // 捲動回表單頂部
      form.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }
});