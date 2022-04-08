<template>
  <header>
    <!-- Navbar -->
    <MDBNavbar  expand="lg" light style="background-color: #e3f2fd"  position="top" container>    
      <MDBNavbarToggler data-toggle="collapse"   target="#navbarExample01"></MDBNavbarToggler>
      <MDBNavbarNav 
      center class="mb-2 mb-lg-0 align-items-center d-flex justify-content-evenly">
        <MDBNavbarBrand href="/" > <i class="fas fa-tint"></i> &nbsp; CUMEX</MDBNavbarBrand>
        
     <!-- Navbar dropdown -->
      <MDBDropdown  class="shadow-0 " btnGroup  v-model="dropdown1">
        <MDBBtn class="shadow-0  nav-link"
          style="background-color: #e3f2fd"
          size="lg"
          @click="dropdown1 = false, navega('/cuenca-mexico')">
          Cuenca</MDBBtn>
        <MDBDropdownToggle tag="a" class="nav-link shadow-0" split @click="dropdown1 = !dropdown1"/>
        <MDBDropdownMenu>
          <MDBDropdownItem href="#" @click="dropdown1 = !dropdown1, navega('/cuenca-hidrologia')">Hidrología</MDBDropdownItem>
          <MDBDropdownItem href="#" @click="dropdown1 = !dropdown1, navega('/cuenca-geologia')">Geología</MDBDropdownItem>
          <MDBDropdownItem href="#" @click="dropdown1 = !dropdown1, navega('/cuenca-vegetacion')">Vegetación</MDBDropdownItem>
          <MDBDropdownItem href="#" @click="dropdown1 = !dropdown1, navega('/cuenca-conservacion')">Conservación</MDBDropdownItem>
          <MDBDropdownItem href="#" @click="dropdown1 = !dropdown1, navega('/cuenca-jurisdiccion')">Marco jurídico</MDBDropdownItem>
        </MDBDropdownMenu>
      </MDBDropdown>

      <MDBDropdown  class="shadow-0 " btnGroup  v-model="dropdown2">
        <MDBBtn class="shadow-0  nav-link"
          style="background-color: #e3f2fd"
          size="lg"
          @click="dropdown2 = false, navega('/cuenca-mexico')">
          Bioindicadores</MDBBtn>
        <MDBDropdownToggle tag="a" class="nav-link shadow-0" split @click="dropdown2 = !dropdown2"/>
        <MDBDropdownMenu>
          <MDBDropdownItem href="#">Calidad fisicoquímica</MDBDropdownItem>
          <MDBDropdownItem href="#">Calidad hidromorfológica</MDBDropdownItem>
          <MDBDropdownItem href="#">Diatomeas</MDBDropdownItem>
          <MDBDropdownItem href="#">Macroalgas</MDBDropdownItem>
          <MDBDropdownItem href="#">Macroinvertebrados</MDBDropdownItem>
        </MDBDropdownMenu>
      </MDBDropdown>

        <MDBBtn class="shadow-0  nav-link"
          style="background-color: #e3f2fd"
          size="lg"
          @click="dropdown1 = false, navega('/cuenca-mexico')">
          Multimedia</MDBBtn>
      </MDBNavbarNav>
    </MDBNavbar>
    <!-- Navbar -->
  </header>
</template>

<script>
  import {
    MDBNavbar,
    MDBNavbarToggler,
    MDBNavbarNav,
    MDBNavbarBrand,
    MDBDropdown,
    MDBDropdownToggle,
    MDBDropdownMenu,
    MDBDropdownItem,
    MDBBtn
  } from 'mdb-vue-ui-kit';
  import { ref } from 'vue';
  import router from '../../router'

  export default {
    components: {
      MDBNavbar,
      MDBNavbarToggler,
      MDBNavbarNav,
      MDBNavbarBrand,
      MDBDropdown,
      MDBDropdownToggle,
      MDBDropdownMenu,
      MDBDropdownItem,
      MDBBtn
    },
    setup() {
      const dropdown1 = ref(false);
      const dropdown2 = ref(false);
      return {
        dropdown1,
        dropdown2,
      };
    },

    methods: {
    navega: function(url) {
      router.push(url).catch(()=>{});
      },
    }
  };
</script>
