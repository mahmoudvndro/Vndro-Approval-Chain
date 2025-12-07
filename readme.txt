If you need to make الحد الاقصي visable again remove that block from frontend.


/* Hide "Max" column text on the Order page only (keeps the column & values) */
#orderPage .products-section table th:nth-child(3),
#orderPage .products-section table td:nth-child(3),
#orderPage .products-section table td:nth-child(3) span {
  color: transparent !important;   /* invisible text */
  text-shadow: none !important;    /* remove any outline/blur */
  user-select: none;               /* optional: prevent copy/paste */
}
