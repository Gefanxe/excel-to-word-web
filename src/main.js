import { createApp } from 'vue';
import ElementPlus from 'element-plus';
import * as ElementPlusIconsVue from '@element-plus/icons-vue';
import 'element-plus/dist/index.css';
import 'virtual:uno.css';
// import './style.css'
import btn from './directives/btn';
import App from './App.vue'

const app = createApp(App);

// 全域註冊 element plus icons
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
  app.component(key, component)
}

app.use(btn);
app.use(ElementPlus);
app.mount('#app');
