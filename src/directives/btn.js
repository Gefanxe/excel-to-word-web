export default {
  install: (app) => {
    const fun = function (evt) {
      let target = evt.target;
      if (target.nodeName === 'SPAN') {
        target = evt.target.parentNode;
      }
      target.blur();
    };
    app.directive('blur', {
      mounted(el) {
        el.addEventListener('click', fun);
      },
      unmounted(el) {
        el.removeEventListener('click', fun);
      }
    });
  }
}
