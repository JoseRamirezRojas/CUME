<template>
  <!-- <img alt="Vue logo" src="./assets/logo.png"> -->
  <div id="app">
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
  }
}
</script>

<style>
html, body {
  height:100%;
  background-color: #fff;
}
#app {
  font-family: Roboto, Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
}
</style>
