<template>
  <header>
    <!-- Navbar -->
    <MDBNavbar  expand="lg" light style="background-color: #e3f2fd"  position="top" container>  
      <!-- Toggle button -->
      <MDBNavbarToggler  target="#navbarExample01" @click="collapse1 = !collapse1"></MDBNavbarToggler>
      <!-- Collapsible wrapper -->
      <MDBCollapse id="navbarExample01" v-model="collapse1">
        <MDBNavbarNav 
        center class="collapse mb-2 mb-lg-0 align-items-center d-flex justify-content-evenly">
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
            @click="dropdown2 = false, navega('/protocolos-info')">
            Bioindicadores</MDBBtn>
          <MDBDropdownToggle tag="a" class="nav-link shadow-0" split @click="dropdown2 = !dropdown2"/>
          <MDBDropdownMenu>
            <MDBDropdownItem href="#" @click="dropdown2 = !dropdown2, navega('/protocolo-fisicoquimica')">Calidad fisicoquímica</MDBDropdownItem>
            <MDBDropdownItem href="#" @click="dropdown2 = !dropdown2, navega('/protocolo-hyqi')">Calidad hidromorfológica</MDBDropdownItem>
            <MDBDropdownItem href="#" @click="dropdown2 = !dropdown2, navega('/protocolo-deqi')">Diatomeas</MDBDropdownItem>
            <MDBDropdownItem href="#" @click="dropdown2 = !dropdown2, navega('/protocolo-meqi')">Macroalgas</MDBDropdownItem>
            <MDBDropdownItem href="#">Macroinvertebrados</MDBDropdownItem>
          </MDBDropdownMenu>
        </MDBDropdown>

          <MDBBtn class="shadow-0  nav-link"
            style="background-color: #e3f2fd"
            size="lg"
            @click="dropdown1 = false, navega('/cuenca-mexico')">
            Multimedia</MDBBtn>
        </MDBNavbarNav>
      </MDBCollapse>
    </MDBNavbar>
    <!-- Navbar -->
  </header>
</template>

<script>
  import {
    MDBNavbar,
    MDBNavbarToggler,
    MDBCollapse,
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
      MDBCollapse,
      MDBNavbarNav,
      MDBNavbarBrand,
      MDBDropdown,
      MDBDropdownToggle,
      MDBDropdownMenu,
      MDBDropdownItem,
      MDBBtn
    },
    setup() {
      const collapse1 = ref(false);
      const dropdown1 = ref(false);
      const dropdown2 = ref(false);
      return {
        collapse1,
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