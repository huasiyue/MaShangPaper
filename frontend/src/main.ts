import { createApp } from "vue";
import { createPinia } from "pinia";
import ArcoVue from "@arco-design/web-vue";
import "@arco-design/web-vue/dist/arco.css";

import App from "./App.vue";
import router from "./router";
import "./styles/arco-theme.less";
import "./styles/global.less";


const app = createApp(App);

app.use(createPinia());
app.use(router);
app.use(ArcoVue);
app.mount("#app");

