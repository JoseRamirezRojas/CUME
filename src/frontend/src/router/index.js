import { createRouter, createWebHistory } from 'vue-router'

import InicioApp            from '@/components/InicioApp'
import CuencaMexico         from '@/components/CuencaMexico'
import CuencaHidrologia     from '@/components/CuencaHidrologia'
import CuencaGeologia       from '@/components/CuencaGeologia'
import CuencaVegetacion     from '@/components/CuencaVegetacion'
import CuencaConservacion   from '@/components/CuencaConservacion'
import CuencaJurisdiccion   from '@/components/CuencaJurisdiccion'
import ProtocolosInicio     from '@/components/ProtocolosInicio'
import ProtocoloFisicoQuim  from '@/components/ProtocoloFisicoQuim'


const routes = [
    // {
    //   path: '/not-found',
    //   name: 'not-found',
    //   component: NotFound
    // },
    {
        path: '/',
        name: 'inicio',
        component: InicioApp
    },
    {
      path: '/cuenca-mexico',
      name: 'cuenca-mexico',
      component: CuencaMexico,  
    },
    {
      path: '/cuenca-hidrologia',
      name: 'cuenca-hidrologia',
      component: CuencaHidrologia,  
    },
    {
      path: '/cuenca-geologia',
      name: 'cuenca-geologia',
      component: CuencaGeologia,  
    },
    {
      path: '/cuenca-vegetacion',
      name: 'cuenca-vegetacion',
      component: CuencaVegetacion,  
    },
    {
      path: '/cuenca-conservacion',
      name: 'cuenca-conservacion',
      component: CuencaConservacion,  
    },
    {
      path: '/cuenca-jurisdiccion',
      name: 'cuenca-jurisdiccion',
      component: CuencaJurisdiccion,  
    },
    {
      path: '/protocolos-info',
      name: 'protocolos-info',
      component: ProtocolosInicio,  
    },
    {
      path: '/protocolo-fisicoquimica',
      name: 'protocolo-fisicoquimica',
      component: ProtocoloFisicoQuim,  
    },
]
const router = createRouter({
  history: createWebHistory(process.env.BASE_URL),
  routes
})
export default router