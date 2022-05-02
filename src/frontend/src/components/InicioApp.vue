<template>
    <!-- Background image -->
    <div
      class="text-center align-items-center d-flex bg-image hover-zoom "
      style="
        height 550px;
      ">
      <img class = "vid" src= '../assets/home_cuenca.jpg '/>
      <div class="mask " style="background-color: rgba(0, 0, 0, 0.3);">
        <div class="d-flex justify-content-center align-items-center h-100">
          <div class="text-white">
            <h1 class="mb-3">Cuenca de México</h1>
            <MDBBtn lg tag="a" outline="light"
              class="m-2"
              role="button"
              @click="navega('/cuenca-mexico')"
              >Más información</MDBBtn>
            <MDBBtn lg tag="a" outline="light"
              class="m-2"
              href="#"
              role="button"
              >Protocolos de evaluación</MDBBtn>
          </div>
        </div>
      </div>
    </div>
    <!-- Background image -->
</template>

<style>
  .vid {
    width:100%;
    display: block;
    margin: auto;
}
</style>

<script>
/* eslint-disable */

  import {
    MDBBtn
  } from 'mdb-vue-ui-kit';
  import { ref } from 'vue';
  import router from '../router'

  export default {
    components: {
      MDBBtn
    },

    methods: {
      navega: function(url) {
      router.push(url).catch(()=>{});
      },
    }
  };
</script>