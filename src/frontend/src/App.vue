<template>
  <!-- <img alt="Vue logo" src="./assets/logo.png"> -->
  <div id="app">
    <button 
        type="button"
        class="btn btn-info btn-floating btn-lg"
        id="btn-back-to-top"
        @click="backToTop()"
        >
  <i class="fas fa-arrow-up"></i>
</button>
    <AppHeader />
    <router-view></router-view>
    <AppFooter />
  </div>
</template>

<script>
import AppHeader from './components/custom/AppHeader.vue'
import AppFooter from './components/custom/AppFooter.vue'

export default {
  name: 'App',
  components: {
    AppHeader,
    AppFooter
  },
  methods: {
    backToTop() {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  }
}

// When the user scrolls down 20px from the top of the document, show the button
window.onscroll = function () {
  scrollFunction();
};

function scrollFunction() {
  if (
    document.body.scrollTop > 40 ||
    document.documentElement.scrollTop > 40
  ) {
    document.getElementById("btn-back-to-top").style.display = "inline";
  } else {
    document.getElementById("btn-back-to-top").style.display = "none";
  }
}

</script>

<style>
html, body {
  height:100%;
  background-color: #fff;
}

#btn-back-to-top {
  position: fixed;
  bottom: 30px;
  right: 30px;
  display: none;
}

#app {
  font-family: Roboto, Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
}
</style>
