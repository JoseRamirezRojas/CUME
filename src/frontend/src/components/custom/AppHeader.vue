<template>
  <header>
    <!-- Navbar -->
    <MDBNavbar expand="lg" light style="background-color: #e3f2fd" container position="sticky">    
      <MDBNavbarToggler target="#navbarExample01"></MDBNavbarToggler>
      <MDBNavbarNav collapse="navbarExample01" 
      center class="mb-2 mb-lg-0 align-items-center d-flex justify-content-evenly">
        <MDBNavbarBrand href="#"> <i class="fas fa-tint"></i> &nbsp; CUMEX</MDBNavbarBrand>
        
     <!-- Navbar dropdown -->
      <MDBDropdown  v-model="dropdown1">
        <MDBDropdownToggle tag="a" class="nav-link" @click="dropdown1 = !dropdown1"
          >Gestión</MDBDropdownToggle>
        <MDBDropdownMenu>
          <MDBDropdownItem href="#">Hidrología</MDBDropdownItem>
          <MDBDropdownItem href="#">Vegetación</MDBDropdownItem>
          <MDBDropdownItem href="#">Conservación</MDBDropdownItem>
          <MDBDropdownItem href="#">Marco jurídico</MDBDropdownItem>
        </MDBDropdownMenu>
      </MDBDropdown>

      <MDBDropdown  v-model="dropdown2">
        <MDBDropdownToggle tag="a" class="nav-link" @click="dropdown2 = !dropdown2"
          >Bioindicadores</MDBDropdownToggle>
        <MDBDropdownMenu>
          <MDBDropdownItem href="#">Calidad fisicoquímica</MDBDropdownItem>
          <MDBDropdownItem href="#">Calidad hidromorfológica</MDBDropdownItem>
          <MDBDropdownItem href="#">Diatomeas</MDBDropdownItem>
          <MDBDropdownItem href="#">Macroalgas</MDBDropdownItem>
          <MDBDropdownItem href="#">Macroinvertebrados</MDBDropdownItem>
        </MDBDropdownMenu>
      </MDBDropdown>

        <MDBNavbarItem href="#">
          Galería
        </MDBNavbarItem>
      </MDBNavbarNav>
    </MDBNavbar>
    <!-- Navbar -->

    <!-- Background image -->
    <div
      class="text-center align-items-center d-flex bg-image hover-zoom"
      style="
        height: 400px;
        margin-top: 0;
      ">
      <img src="https://mdbootstrap.com/img/new/slides/041.webp"/>
      <div class="mask " style="background-color: rgba(0, 0, 0, 0.6);">
        <div class="d-flex justify-content-center align-items-center h-100">
          <div class="text-white">
            <h1 class="mb-3">Cuenca de México</h1>
            <MDBBtn lg tag="a" outline="light"
              class="m-2"
              href="#"
              role="button"
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
  </header>
</template>

<script>
  import {
    MDBNavbar,
    MDBBtn,
    MDBNavbarToggler,
    MDBNavbarNav,
    MDBNavbarBrand,
    MDBNavbarItem,
    MDBDropdown,
    MDBDropdownToggle,
    MDBDropdownMenu,
    MDBDropdownItem,
  } from 'mdb-vue-ui-kit';
  import { ref } from 'vue';
  export default {
    components: {
      MDBNavbar,
      MDBBtn,
      MDBNavbarToggler,
      MDBNavbarNav,
      MDBNavbarBrand,
      MDBNavbarItem,
      MDBDropdown,
      MDBDropdownToggle,
      MDBDropdownMenu,
      MDBDropdownItem,
    },
    setup() {
      const dropdown2 = ref(false);
      const dropdown1 = ref(false);
      return {
        dropdown1,
        dropdown2,
      };
    },
  };
</script>



