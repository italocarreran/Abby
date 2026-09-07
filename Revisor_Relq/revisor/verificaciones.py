# -*- coding: utf-8 -*-
"""Motor de las comprobaciones V4…V17 del Revisor.

No importa el punto de entrada para evitar ciclos. ``configurar`` recibe una vez
las constantes y adaptadores históricos que usa el motor; la ventana solo llama
a ``comprobar`` y conserva el orden observable de ejecución.
"""

_dependencias = None

def configurar(dependencias):
    """Inyecta el espacio histórico del Revisor sin importar su ventana."""
    global _dependencias
    _dependencias = dependencias
    globals().update({k: v for k, v in dependencias.items()
                      if not k.startswith("__") and k not in globals()})

def comprobar(self, c, valores, L):
    """Corre una comprobacion y devuelve un dict con su resultado."""
    tipo = c["tipo"]
    base = {"tipo": tipo, "desc": c.get("desc", tipo)}

    if tipo == "igualdad":
        faltan = [k for k in c["izq"] + c["der"]
                  if valores.get(partir_signo(k)[0]) is None]
        if faltan:
            L(f"  ? {c['desc']}: sin datos ({', '.join(faltan)})")
            return dict(base, estado="SIN DATOS")
        izq = sum(valores[k] * sg for k, sg in map(partir_signo, c["izq"]))
        der = sum(valores[k] * sg for k, sg in map(partir_signo, c["der"]))
        absoluto = bool(c.get("absoluto"))
        if absoluto:
            izq, der = abs(izq), abs(der)
        dif = izq - der
        ok = abs(dif) <= TOLERANCIA
        L(f"  {'OK ' if ok else '>> '}{c['desc']}"
          + ("   (en valor absoluto)" if absoluto else ""))
        L(f"        {fmt_monto(izq)}  vs  {fmt_monto(der)}"
          f"   dif {fmt_monto(dif)}")
        # Si al cambiarle el signo a un lado la diferencia se hace mucho mas
        # chica, lo que falla es el signo y no los montos. Se avisa para no
        # dejar a la vista un descuadre gigante que en realidad no lo es.
        suma = izq + der
        if not ok and not absoluto and abs(suma) < abs(dif) / 2:
            L(f"        OJO: parece un problema de signo. Con un lado "
              f"invertido la diferencia sería {fmt_monto(suma)}.")
            if abs(suma) <= TOLERANCIA:
                L("        Los dos montos son iguales y de signo opuesto: "
                  "seguramente falta un '-' delante de una clave.")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    izquierda=izq, derecha=der, diferencia=dif,
                    dif_invertida=suma, absoluto=absoluto,
                    izq_claves=list(c["izq"]), der_claves=list(c["der"]))

    if tipo == "umbral":
        # |valor| <= maximo. Para residuos de redondeo, donde exigir 0 seria
        # irreal pero un numero grande delata un problema de verdad.
        clave = c["clave"]
        val = valores.get(clave)
        if val is None:
            L(f"  ? {c['desc']}: sin datos ({clave})")
            return dict(base, estado="SIN DATOS")
        tope = c.get("maximo", UMBRAL_DESCUADRE_CPRT)
        ok = abs(val) <= tope
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {VALORES[clave]['etiqueta']}")
        L(f"        vale {fmt_monto(val)}   (máximo aceptado {fmt_monto(tope)})")
        if not ok:
            L("        Un descuadre grande suele ser que el Cuadro de pagos se armó")
            L("        con una tabla distinta de la que quedó, o que faltó apretar")
            L("        «Actualiza Rango» antes de refrescar la tabla dinámica.")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    valor=val, maximo=tope)

    if tipo == "cero":
        faltan = [k for k in c["claves"]
                  if valores.get(partir_signo(k)[0]) is None]
        if faltan:
            L(f"  ? {c['desc']}: sin datos ({', '.join(faltan)})")
            return dict(base, estado="SIN DATOS")
        malos = {}
        for k in c["claves"]:
            bk, sg = partir_signo(k)
            if abs(valores[bk] * sg) > TOLERANCIA:
                malos[k] = valores[bk] * sg
        ok = not malos
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        for k, val in malos.items():
            L(f"        {VALORES[partir_signo(k)[0]]['etiqueta']} debería ser 0 "
              f"y vale {fmt_monto(val)}")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    valores={k: valores[partir_signo(k)[0]] * partir_signo(k)[1]
                             for k in c["claves"]},
                    fuera=list(malos))

    if tipo == "marcas":
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        res = buscar_marcas_rapido(ruta, c["hoja"], c["fila_inicio"],
                                   c["reglas"], L)
        if res is None:
            return dict(base, estado="SIN DATOS")
        ok = not res["conteo"]
        rangos = ", ".join(r for regla in c["reglas"] for r in regla["rangos"])
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {rangos}, desde la fila {c['fila_inicio']}")
        if ok:
            L("        sin errores ni textos prohibidos")
        for motivo, n in res["conteo"].items():
            L(f"        {n} celda(s) con {motivo}")
        for celda, motivo, val in res["marcas"]:
            L(f"          {celda:<9} {motivo:<18} {val}")
        total = sum(res["conteo"].values())
        if total > len(res["marcas"]):
            L(f"          ... y {total - len(res['marcas'])} más")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    conteo=res["conteo"], marcas=res["marcas"])

    if tipo == "tabla":
        ra = self.rutas.get(c["archivo_a"])
        rb = self.rutas.get(c["archivo_b"])
        if ra is None or rb is None:
            L(f"  ? {c['desc']}: falta uno de los archivos")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        da = leer_columnas_rapido(ra, c["hoja_a"], c["cols_a"], c["fila_a"], L)
        db = leer_columnas_rapido(rb, c["hoja_b"], c["cols_b"], c["fila_b"], L)
        if da is None or db is None:
            return dict(base, estado="SIN DATOS")
        info_a, info_b = {}, {}
        ta = armar_tabla(da, c["cols_a"][0], c["cols_a"][1:], L,
                         c.get("nombre_a", "A"), info=info_a)
        tb = armar_tabla(db, c["cols_b"][0], c["cols_b"][1:], L,
                         c.get("nombre_b", "B"), info=info_b)
        solo_a = [ta[k][0] for k in ta if k not in tb]
        solo_b = [tb[k][0] for k in tb if k not in ta]
        difs = []
        for k in ta:
            if k not in tb:
                continue
            for i, (va, vb) in enumerate(zip(ta[k][1], tb[k][1])):
                na = va if isinstance(va, (int, float)) else None
                nb = vb if isinstance(vb, (int, float)) else None
                if na is None and nb is None:
                    continue
                if na is None or nb is None or abs(na - nb) > TOLERANCIA:
                    difs.append((ta[k][0], i, va, vb))
        # Las empresas repetidas son un fallo, no un aviso: la tabla se pega
        # copiada y una empresa dos veces significa que a alguien se le paga
        # o se le cobra dos veces.
        dups = []
        if c.get("sin_duplicados"):
            dups = ([(c.get("nombre_a", "A"), n, f) for n, f in info_a["duplicadas"]]
                    + [(c.get("nombre_b", "B"), n, f) for n, f in info_b["duplicadas"]])
        ok = not (solo_a or solo_b or difs or dups)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        for et, n, f in dups[:10]:
            L(f"        REPETIDA en {et}: {n[:34]} (fila {f})")
        if len(dups) > 10:
            L(f"        ... y {len(dups) - 10} repetición(es) más")
        if solo_a:
            L(f"        {len(solo_a)} empresa(s) solo en {c.get('nombre_a')}: "
              f"{', '.join(sorted(solo_a)[:10])}")
        if solo_b:
            L(f"        {len(solo_b)} empresa(s) solo en {c.get('nombre_b')}: "
              f"{', '.join(sorted(solo_b)[:10])}")
        if difs:
            nombres = c.get("nombres_valor") or ["valor 1", "valor 2"]
            L(f"        {len(difs)} diferencia(s) de monto:")
            for emp, i, va, vb in difs[:20]:
                et = nombres[i] if i < len(nombres) else f"col {i + 2}"
                dif = (va - vb) if isinstance(va, (int, float)) and isinstance(vb, (int, float)) else None
                L(f"          {emp[:28]:<30} {et:<12} {fmt_monto(va):>18} vs "
                  f"{fmt_monto(vb):>18}   dif {fmt_monto(dif)}")
            if len(difs) > 20:
                L(f"          ... y {len(difs) - 20} más")
        if ok:
            L(f"        las {len(ta)} empresas y sus montos coinciden")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    solo_a=solo_a, solo_b=solo_b, n_difs=len(difs),
                    difs=[(e, i, va, vb) for e, i, va, vb in difs[:40]],
                    duplicadas=[(et, n, f) for et, n, f in dups[:40]],
                    filas_a=len(ta), filas_b=len(tb))

    if tipo == "formulas_cubren":
        # Las formulas de unas columnas tienen que llegar exactamente hasta
        # la ultima empresa: ni cortarse antes (empresa sin calcular) ni
        # seguir despues (arrastra 0 o vacio y ensucia los totales).
        ruta = self.rutas.get(c["archivo"])
        ref = c["referencia"]
        ruta_ref = self.rutas.get(ref["archivo"])
        if ruta is None or ruta_ref is None:
            L(f"  ? {c['desc']}: falta uno de los archivos")
            return dict(base, estado="SIN DATOS")
        cols = []
        for r in c["cols"]:
            cols.extend(expandir_columnas(r))
        f_ini = int(c["fila_inicio"])
        f_ini_ref = int(ref.get("fila_inicio", 1))

        datos_ref = leer_columnas_rapido(ruta_ref, ref["hoja"], [ref["col"]],
                                         f_ini_ref, L)
        if datos_ref is None:
            return dict(base, estado="SIN DATOS")
        # Ultima fila con una empresa de verdad en la columna de referencia.
        filas_emp = [f for f, v in datos_ref.get(ref["col"], {}).items()
                     if isinstance(v, str) and v.strip()
                     and normalizar(v) not in ("", "0")
                     and not v.startswith("#")]
        if not filas_emp:
            L(f"  ? {c['desc']}: no hay empresas en "
              f"{ref['col']}{f_ini_ref} hacia abajo")
            return dict(base, estado="SIN DATOS")
        ultima_ref = max(filas_emp)
        n_emp = len(filas_emp)
        # Las filas pueden estar corridas entre hojas (K desde 5, A desde 9).
        desfase = f_ini - f_ini_ref
        esperada = ultima_ref + desfase

        formulas = leer_formulas_rapido(ruta, c["hoja"], cols, f_ini, L)
        if formulas is None:
            return dict(base, estado="SIN DATOS")
        # Tambien se mira hasta donde hay VALOR, no solo formula. Motivos:
        #  - un 0 arrastrado por una formula que sobra es un valor, y hay que
        #    cazarlo aunque la formula este "bien" puesta;
        #  - si la columna trae valores pegados a mano en vez de formulas,
        #    igual hay que revisar que no sobren filas.
        valores_col = leer_columnas_rapido(ruta, c["hoja"], cols, f_ini, L)
        if valores_col is None:
            valores_col = {}

        faltan, sobran, sin_formula = [], [], []
        for col in cols:
            fs = formulas.get(col, set())
            vs = set(valores_col.get(col, {}))
            if not fs:
                sin_formula.append(col)
            alcance = max(fs | vs) if (fs or vs) else None
            if alcance is None:
                faltan.append((col, 0, esperada - f_ini + 1))
                continue
            if alcance < esperada:
                faltan.append((col, alcance, esperada - alcance))
            elif alcance > esperada:
                sobran.append((col, alcance, alcance - esperada))
        # Que una columna traiga valores pegados en vez de formulas no es un
        # error por si mismo: se avisa, pero lo que decide es el alcance.
        # Con solo_faltan, que las formulas sigan mas abajo NO es un error.
        # Hace falta para la H del CPRT: ahi el bloque de datos lo genera una
        # tabla dinamica que crece y se encoge, la formula esta arrastrada bien
        # abajo a proposito, y el exportador corta por la columna A. Lo unico
        # que importa es que no se quede corta.
        if c.get("solo_faltan"):
            sobran = []
        ok = not (faltan or sobran)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {n_emp} empresa(s) en {ref['col']}{f_ini_ref}:"
          f"{ref['col']}{ultima_ref}"
          + (f"   (las fórmulas van corridas {desfase:+d} fila(s))" if desfase else ""))
        L(f"        {'+'.join(c['cols'])} debería llegar hasta la fila {esperada}")
        for col, ult, n in faltan:
            L(f"          {col}: llega a la fila {ult}, FALTAN {n} "
              f"-> hay empresas sin calcular")
        for col, ult, n in sobran:
            L(f"          {col}: llega a la fila {ult}, SOBRAN {n} "
              f"-> arrastra 0 o vacío por debajo de la última empresa")
        if sin_formula:
            L(f"        (sin fórmulas, con valores escritos: "
              f"{', '.join(sin_formula)})")
        if ok:
            L(f"        las {len(cols)} columna(s) llegan justo a la fila {esperada}")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    esperada=esperada, n_empresas=n_emp,
                    faltan=faltan[:40], sobran=sobran[:40],
                    sin_formula=sin_formula)

    if tipo == "suma_por_empresa":
        # La columna de detalle trae la empresa repetida (una fila por
        # reemplazo). Se suma por empresa y se compara contra el total de la
        # tabla resumen. Todo dentro del mismo archivo y la misma hoja.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        cols_res = list(c["cols_resumen"])          # [empresa, val1, val2, ...]
        necesarias = [c["col_empresa"], c["col_monto"]] + cols_res
        datos = leer_columnas_rapido(ruta, c["hoja"], necesarias,
                                     c["fila_inicio"], L)
        if datos is None:
            return dict(base, estado="SIN DATOS")

        # 1) detalle: sumar el monto agrupando por empresa
        detalle, sin_empresa = {}, 0
        col_e, col_m = c["col_empresa"], c["col_monto"]
        for f in sorted(datos.get(col_e, {})):
            nombre = datos[col_e][f]
            if not isinstance(nombre, str) or not nombre.strip():
                continue
            k = normalizar(nombre)
            if k in ("", "0"):
                continue
            v = datos.get(col_m, {}).get(f)
            if not isinstance(v, (int, float)):
                if v is not None:
                    sin_empresa += 1
                continue
            acum, nom, filas = detalle.get(k, (0.0, nombre.strip(), 0))
            detalle[k] = (acum + float(v), nom, filas + 1)

        # 2) resumen: la tabla de totales
        info = {}
        resumen = armar_tabla(datos, cols_res[0], cols_res[1:], L,
                              c.get("nombre_resumen", "resumen"), info=info)

        difs, solo_det, solo_res = [], [], []
        for k, (suma, nom, nfilas) in detalle.items():
            if k not in resumen:
                solo_det.append(nom)
                continue
            vals = [v for v in resumen[k][1] if isinstance(v, (int, float))]
            total = sum(vals)
            if abs(suma - total) > TOLERANCIA:
                difs.append((nom, suma, total, suma - total, nfilas))
        for k in resumen:
            if k not in detalle:
                solo_res.append(resumen[k][0])

        dups = [n for n, _ in info.get("duplicadas", [])]
        ok = not (difs or solo_det or solo_res or dups)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {len(detalle)} empresa(s) distintas en {col_e} "
          f"(sumando {col_m}), {len(resumen)} en la tabla resumen")
        if sin_empresa:
            L(f"        {sin_empresa} fila(s) con empresa pero sin monto numérico")
        for n in solo_det[:15]:
            L(f"          en el detalle y NO en el resumen: {n[:40]}")
        for n in solo_res[:15]:
            L(f"          en el resumen y NO en el detalle: {n[:40]}")
        for n in dups[:10]:
            L(f"          REPETIDA en el resumen: {n[:40]}")
        if difs:
            etiqueta = " + ".join(cols_res[1:])
            L(f"        {len(difs)} empresa(s) descuadrada(s) "
              f"(suma de {col_m}  vs  {etiqueta}):")
            for nom, suma, total, dif, nfilas in sorted(
                    difs, key=lambda x: -abs(x[3]))[:20]:
                L(f"          {nom[:26]:<28} {fmt_monto(suma):>16} vs "
                  f"{fmt_monto(total):>16}  dif {fmt_monto(dif):>14}"
                  f"  ({nfilas} fila(s))")
            if len(difs) > 20:
                L(f"          ... y {len(difs) - 20} más")
            # Si invirtiendo el signo cuadra, el problema es de convencion.
            invertidos = sum(1 for _, s, t, _, _ in difs
                             if abs(s + t) <= TOLERANCIA)
            if invertidos:
                L(f"        OJO: en {invertidos} de esas empresas cuadraría "
                  f"con el signo invertido. Puede que el neto se calcule "
                  f"como recibe - paga y no como paga + recibe.")
        if ok:
            L(f"        las {len(detalle)} empresas cuadran una por una")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_detalle=len(detalle), n_resumen=len(resumen),
                    difs=[(n, s, t, d, nf) for n, s, t, d, nf in
                          sorted(difs, key=lambda x: -abs(x[3]))[:40]],
                    n_difs=len(difs), solo_detalle=solo_det[:40],
                    solo_resumen=solo_res[:40], duplicadas=dups[:40])

    if tipo == "suma_calculada":
        # Recalcula un monto FILA POR FILA a partir de sus componentes y lo
        # compara contra la columna que ya lo trae calculado.
        #   valor_fila = (suma de "terminos", con su signo) * (los "factores")
        # En el tabulado: (CV - CMg) * Generacion * USD  contra la columna E.
        #
        # Se compara el TOTAL, que es lo que se pidio, pero tambien se cuenta
        # cuantas filas fallan por su cuenta: en un total, dos errores de
        # signo contrario se cancelan y no se ven. Y saber si falla UNA fila o
        # TODAS distingue un dato malo de un problema de redondeo.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")

        terminos = [(t.lstrip("+-").upper(), -1.0 if t.startswith("-") else 1.0)
                    for t in c["terminos"]]
        factores = [f.upper() for f in c.get("factores", [])]
        contra = c["contra"].upper()
        cols = sorted({t for t, _ in terminos} | set(factores) | {contra})
        f_ini = int(c["fila_inicio"])

        datos = leer_columnas_rapido(ruta, c["hoja"], cols, f_ini, L)
        if datos is None:
            return dict(base, estado="SIN DATOS")

        def num(col, fila):
            v = datos.get(col, {}).get(fila)
            return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        filas = sorted({f for col in cols for f in datos.get(col, {})
                        if f >= f_ini})
        tol_fila = float(c.get("tolerancia_fila", 0.5))
        suma_calc = suma_real = 0.0
        n_usadas = n_sin_datos = 0
        malas = []
        for f in filas:
            comps = {col: num(col, f) for col in cols}
            real = comps[contra]
            # Una fila sin ningun dato es el relleno del final: se ignora.
            if real is None and all(comps[col] is None for col, _ in terminos):
                continue
            if any(comps[col] is None for col, _ in terminos) or \
               any(comps[col] is None for col in factores):
                n_sin_datos += 1
                continue
            calc = sum(comps[col] * sg for col, sg in terminos)
            for col in factores:
                calc *= comps[col]
            real = real if real is not None else 0.0
            suma_calc += calc
            suma_real += real
            n_usadas += 1
            if abs(calc - real) > tol_fila:
                malas.append((f, calc, real, calc - real))

        if not n_usadas:
            L(f"  ? {c['desc']}: no hubo ninguna fila con todos los datos")
            return dict(base, estado="SIN DATOS")

        formula = " ".join(
            (("- " if sg < 0 else ("+ " if i else "")) + col)
            for i, (col, sg) in enumerate(terminos))
        if factores:
            formula = f"({formula}) * " + " * ".join(factores)
        dif = suma_calc - suma_real
        tol = float(c.get("tolerancia", TOLERANCIA))
        ok = abs(dif) <= tol
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {n_usadas} fila(s) usadas, desde la {f_ini}"
          + (f"; {n_sin_datos} omitida(s) por falta de algún componente"
             if n_sin_datos else ""))
        L(f"        {formula}  =  {fmt_monto(suma_calc)}")
        L(f"        columna {contra}          =  {fmt_monto(suma_real)}")
        L(f"        diferencia            =  {fmt_monto(dif)}"
          f"   (máximo aceptado {fmt_monto(tol)})")
        if suma_real:
            L(f"        en proporción         =  {dif / suma_real:.2%}")
        if malas:
            L(f"        {len(malas)} de {n_usadas} fila(s) no cuadran por su "
              f"cuenta (más de {fmt_monto(tol_fila)} de diferencia):")
            for f, calc, real, d in sorted(malas, key=lambda x: -abs(x[3]))[:12]:
                L(f"          fila {f}: calculado {fmt_monto(calc)} vs "
                  f"{contra}{f} {fmt_monto(real)}   dif {fmt_monto(d)}")
            if len(malas) > 12:
                L(f"          ... y {len(malas) - 12} fila(s) más")
            if len(malas) == n_usadas:
                L("        Fallan TODAS las filas: mirá si es redondeo o si "
                  "alguna columna no es la que se cree.")
        elif ok:
            L(f"        las {n_usadas} filas cuadran una por una")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    suma_calculada=suma_calc, suma_real=suma_real,
                    diferencia=dif, tolerancia=tol, formula=formula,
                    n_filas=n_usadas, n_sin_datos=n_sin_datos,
                    n_malas=len(malas),
                    malas=[(f, a, b, d) for f, a, b, d in
                           sorted(malas, key=lambda x: -abs(x[3]))[:40]])

    if tipo == "retencion_coherente":
        # La H del CPRT es la G con las retenciones puestas en 0:
        #     H = SI(CONTAR.SI(retenciones; acreedor)<>0; 0; 1) * G
        # O sea que en cada fila la H solo puede ser IGUAL a la G, o CERO.
        # Cualquier otro valor significa que la formula de la H esta mal, y eso
        # SI es un error, porque el csv se arma con la H.
        # Que haya filas en 0 no es un error: es una retencion de verdad. Pero
        # se avisa con el monto, porque es plata que no se paga.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        cg, ch, cref = c["col_g"], c["col_h"], c["col_referencia"]
        d = leer_columnas_rapido(ruta, c["hoja"], [cg, ch, cref],
                                 c["fila_inicio"], L)
        if d is None:
            return dict(base, estado="SIN DATOS")
        filas = [f for f, v in d.get(cref.upper(), {}).items()
                 if isinstance(v, str) and v.strip()]
        iguales, retenidas, raras, sin_h = 0, [], [], []
        total_ret = 0.0
        for f in sorted(filas):
            g = d.get(cg.upper(), {}).get(f)
            h = d.get(ch.upper(), {}).get(f)
            if not isinstance(g, (int, float)) or isinstance(g, bool):
                continue
            if not isinstance(h, (int, float)) or isinstance(h, bool):
                sin_h.append(f)      # la formula de la H no llega hasta aca
                continue
            if abs(h - g) <= TOLERANCIA:
                iguales += 1
            elif abs(h) <= TOLERANCIA:
                retenidas.append((f, float(g)))
                total_ret += float(g)
            else:
                raras.append((f, float(g), float(h)))
        ok = not raras and not sin_h
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {len(filas)} fila(s): {iguales} con {ch} = {cg}, "
          f"{len(retenidas)} retenida(s) en 0")
        if sin_h:
            L(f"        {len(sin_h)} fila(s) SIN valor en {ch}: la fórmula no "
              f"llega hasta abajo. El csv saldría con el monto vacío.")
            for f in sin_h[:10]:
                L(f"             fila {f}")
        if raras:
            L(f"        {len(raras)} fila(s) con {ch} que no es {cg} ni 0: "
              f"la fórmula de la {ch} está mal.")
            for f, g, h in sorted(raras, key=lambda x: -abs(x[1] - x[2]))[:10]:
                L(f"             fila {f}: {cg}={fmt_monto(g)}  "
                  f"{ch}={fmt_monto(h)}")
        if retenidas:
            L(f"        OJO: {fmt_monto(total_ret)} retenido en total. El csv "
              f"se arma con la {ch}, así que esas empresas NO reciben pago:")
            for f, g in sorted(retenidas, key=lambda x: -x[1])[:10]:
                L(f"             fila {f}: {fmt_monto(g)}")
            if len(retenidas) > 10:
                L(f"             ... y {len(retenidas) - 10} más")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_filas=len(filas), n_iguales=iguales,
                    n_retenidas=len(retenidas), total_retenido=total_ret,
                    retenidas=[(f, g) for f, g in retenidas[:40]],
                    raras=[(f, g, h) for f, g, h in raras[:40]],
                    sin_h=sin_h[:40])

    if tipo == "prorrata_al_dia":
        # La hoja PRORRATA_RETIROS de las planillas 3, 5 y 6 es el pivote de
        # la tabla larga del Prorrata_Retiros: mismos numeros, otra forma.
        #     origen  (PRORRATA_HORARIA_TABULAR, desde la fila 2):
        #             A Hora | B Suministrador | C Prorrata_horaria
        #     destino (PRORRATA_RETIROS): B8="Hora", C8.. suministradores,
        #             B9.. las horas, el resto los valores (0 si falta)
        #
        # Como son los MISMOS numeros reordenados, la suma por suministrador
        # tiene que coincidir exactamente. Eso caza el olvido de actualizar:
        # otro mes trae otros totales y casi siempre otros suministradores.
        r_dest = self.rutas.get(c["archivo"])
        r_orig = self.rutas.get(c["origen"]["archivo"])
        if r_dest is None or r_orig is None:
            L(f"  ? {c['desc']}: falta uno de los archivos")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")

        # --- origen: tabla larga -> suma por suministrador ---
        o = c["origen"]
        do = leer_columnas_rapido(
            r_orig, o["hoja"], [o["col_hora"], o["col_suministrador"],
                                o["col_valor"]], o["fila_inicio"], L)
        if do is None:
            return dict(base, estado="SIN DATOS")
        ch, cs, cv = (x.upper() for x in (o["col_hora"],
                                          o["col_suministrador"],
                                          o["col_valor"]))
        suma_orig, horas_orig, n_orig = {}, set(), 0
        for f in sorted(do.get(cs, {})):
            sumi = do[cs].get(f)
            if not isinstance(sumi, str) or not sumi.strip():
                continue
            v = do.get(cv, {}).get(f)
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                continue
            k = normalizar(sumi)
            suma_orig[k] = suma_orig.get(k, 0.0) + float(v)
            h = do.get(ch, {}).get(f)
            if h is not None:
                horas_orig.add(str(h).strip())
            n_orig += 1
        L(f"        origen : {n_orig} fila(s), {len(suma_orig)} "
          f"suministrador(es), {len(horas_orig)} hora(s)")

        # --- destino: la matriz ---
        fila_enc = int(c["fila_encabezado"])
        cols = columnas_de_fila(r_dest, c["hoja"], fila_enc, L)
        if cols is None:
            return dict(base, estado="SIN DATOS")
        n_ini = col_letra_a_num(c["col_inicio"])
        cols = [x for x in cols if col_letra_a_num(x) >= n_ini]
        if not cols:
            L(f"  >> {c['desc']}")
            L(f"        la fila {fila_enc} de {c['hoja']} está VACÍA desde "
              f"{c['col_inicio']}: la prorrata nunca se pegó.")
            return dict(base, estado="NO CUADRA", n_origen=len(suma_orig),
                        n_destino=0, difs=[], solo_origen=[], solo_destino=[])
        datos = leer_columnas_rapido(r_dest, c["hoja"], cols, fila_enc, L)
        if datos is None:
            return dict(base, estado="SIN DATOS")
        # La primera columna es "Hora"; de la segunda en adelante, un
        # suministrador por columna.
        #
        # OJO con dos cosas de estas hojas:
        #  - A la DERECHA del bloque de prorrata puede haber mas columnas con
        #    los MISMOS nombres de suministrador pero con MONTOS (las
        #    planillas 5 y 6 los tienen), y columnas de total. Por eso se toma
        #    solo la PRIMERA columna de cada nombre: el bloque pegado empieza
        #    en B8, asi que es el de mas a la izquierda. Sumar las dos daba
        #    numeros de millones contra prorratas de dos digitos.
        #  - En la planilla 3 puede haber centrales pegadas a mano que quedan
        #    en 0%. Esas no estan en el origen y no son un error.
        col_hora = cols[0]
        filas_dato = sorted(f for f in datos.get(col_hora, {})
                            if f > fila_enc)
        n_filas_dest = len(filas_dato)
        suma_dest, vistos, repetidas = {}, set(), []
        for col in cols[1:]:
            nom = datos.get(col, {}).get(fila_enc)
            if not isinstance(nom, str) or not nom.strip():
                continue
            k = normalizar(nom)
            if k in vistos:
                repetidas.append(f"{nom.strip()} ({col})")
                continue
            vistos.add(k)
            tot = 0.0
            for f in filas_dato:
                v = datos.get(col, {}).get(f)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    tot += float(v)
            suma_dest[k] = tot
        L(f"        destino: {n_filas_dest} fila(s) de hora, "
          f"{len(suma_dest)} columna(s) con nombre")
        if repetidas:
            L(f"        {len(repetidas)} columna(s) con un nombre ya visto "
              f"(bloques de montos a la derecha): se usa la primera de cada "
              f"una")

        # --- comparacion ---
        # La exigencia es ASIMETRICA a proposito: todo suministrador del
        # ORIGEN tiene que estar en la planilla con la misma suma. Lo que
        # sobra en la planilla solo se informa, porque hay dos motivos
        # legitimos: centrales pegadas a mano que quedan en 0% (planilla 3) y
        # columnas de totales o de montos que son parte de la hoja
        # (planillas 5 y 6).
        tol = c.get("tolerancia", TOL_PRORRATA_SUMA)
        solo_o = sorted(k for k in suma_orig if k not in suma_dest)
        sobran = sorted(k for k in suma_dest if k not in suma_orig)
        # De lo que sobra, solo se nombra lo que tiene algo distinto de cero:
        # una central en 0% no aporta nada y no vale la pena listarla.
        sobran_con_valor = [k for k in sobran if abs(suma_dest[k]) > tol]
        sobran_en_cero = [k for k in sobran if abs(suma_dest[k]) <= tol]
        difs = []
        n_ok = 0
        for k, v in suma_orig.items():
            if k not in suma_dest:
                continue
            d = suma_dest[k] - v
            if abs(d) > tol:
                difs.append((k, suma_dest[k], v, d))
            else:
                n_ok += 1
        # Las horas: solo se avisa. La hoja puede tener filas de mas abajo
        # que no son parte de la matriz.
        dif_horas = bool(horas_orig) and len(horas_orig) != n_filas_dest
        ok = not (difs or solo_o)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {n_ok} de {len(suma_orig)} suministrador(es) del origen "
          f"cuadran")
        if dif_horas:
            L(f"        (las horas no coinciden: {len(horas_orig)} en el "
              f"origen y {n_filas_dest} filas en la planilla)")
        for k in solo_o[:10]:
            L(f"          FALTA en la planilla: {k[:40]} "
              f"(origen {suma_orig[k]:.6f})")
        if len(solo_o) > 10:
            L(f"          ... y {len(solo_o) - 10} más")
        for k, vd, vo, d in sorted(difs, key=lambda x: -abs(x[3]))[:10]:
            L(f"          {k[:34]:<36} planilla {vd:.6f} vs origen {vo:.6f}")
        if len(difs) > 10:
            L(f"          ... y {len(difs) - 10} más")
        if sobran_en_cero:
            L(f"        {len(sobran_en_cero)} columna(s) de la planilla en 0 "
              f"que no están en el origen (normal): "
              + ", ".join(x[:20] for x in sobran_en_cero[:6])
              + (" ..." if len(sobran_en_cero) > 6 else ""))
        if sobran_con_valor:
            L(f"        {len(sobran_con_valor)} columna(s) de la planilla que "
              f"no son del origen (totales, montos): "
              + ", ".join(x[:22] for x in sobran_con_valor[:6])
              + (" ..." if len(sobran_con_valor) > 6 else ""))
        if not ok:
            L("        La prorrata de esta planilla NO es la del "
              "Prorrata_Retiros de este mes.")
            L("        Hay que correr «Actualizar data» y marcar Prorrata.")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_origen=len(suma_orig), n_destino=len(suma_dest),
                    n_ok=n_ok, horas_origen=len(horas_orig),
                    horas_destino=n_filas_dest,
                    solo_origen=solo_o[:40],
                    solo_destino=sobran_con_valor[:40],
                    sobran_en_cero=len(sobran_en_cero),
                    difs=[(k, vd, vo, d) for k, vd, vo, d in
                          sorted(difs, key=lambda x: -abs(x[3]))[:40]])

    if tipo == "centrales_sin_dueno":
        # Cruza dos tablas del mismo Access: toda central con MONTO en
        # Sobrecostos tiene que tener dueño en Central_Empresa.
        #
        # Una central sin dueño NO es un error por si misma: en
        # CONSUMOS_PROPIOS de la planilla 6 hay centrales sin propietario a
        # proposito. Lo que no puede pasar es que una central con plata no
        # tenga a quien pagarle o a quien cobrarle.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        t_sob = c.get("tabla_montos", "Sobrecostos")
        t_ce = c.get("tabla_duenos", "Central_Empresa")
        cn = None
        try:
            cn = conexion_mdb(ruta)
            cur = cn.cursor()
            tablas = {r.table_name for r in cur.tables(tableType="TABLE")}
            for t in (t_sob, t_ce):
                if t not in tablas:
                    L(f"  ? {c['desc']}: no está la tabla [{t}]. "
                      f"Hay: {', '.join(sorted(tablas))}")
                    return dict(base, estado="SIN DATOS")

            def col(tabla, objetivo):
                cols = [r.column_name for r in cur.columns(table=tabla)]
                for x in cols:
                    if normalizar(x) == normalizar(objetivo):
                        return x
                for x in cols:
                    if normalizar(objetivo) in normalizar(x):
                        return x
                return None

            c_cen_s = col(t_sob, "Central")
            c_mon = col(t_sob, "Sobrecosto")
            c_cen_e = col(t_ce, "Central")
            c_emp = col(t_ce, "Empresa")
            if not all((c_cen_s, c_mon, c_cen_e, c_emp)):
                L(f"  ? {c['desc']}: no se encontraron las columnas "
                  f"necesarias en las dos tablas.")
                return dict(base, estado="SIN DATOS")

            # Dueños: por central, si tiene una empresa de verdad.
            duenos = {}
            cur.execute(f"SELECT [{c_cen_e}], [{c_emp}] FROM [{t_ce}]")
            for cen, emp in cur.fetchall():
                if cen is None or not str(cen).strip():
                    continue
                k = clave_central(cen)
                tiene = (emp is not None and str(emp).strip() != ""
                         and str(emp).strip() != "0")
                duenos[k] = duenos.get(k, False) or tiene

            cur.execute(f"SELECT [{c_cen_s}], SUM([{c_mon}]) FROM [{t_sob}] "
                        f"GROUP BY [{c_cen_s}]")
            montos = []
            for cen, suma in cur.fetchall():
                if cen is None or not str(cen).strip():
                    montos.append((None, "(central vacía)", float(suma or 0)))
                else:
                    montos.append((clave_central(cen), str(cen).strip(),
                                   float(suma or 0)))
        except Exception as e:
            L(f"  ? {c['desc']}: no se pudo leer el Access: {e}")
            return dict(base, estado="SIN DATOS")
        finally:
            try:
                if cn is not None:
                    cn.close()
            except Exception:
                pass

        tol = c.get("tolerancia", TOLERANCIA)
        sin_dueno, sin_fila, con_dueno, sin_plata = [], [], 0, 0
        for k, nombre, suma in montos:
            if abs(suma) <= tol:
                sin_plata += 1
            elif k is None or k not in duenos:
                sin_fila.append((nombre, suma))
            elif not duenos[k]:
                sin_dueno.append((nombre, suma))
            else:
                con_dueno += 1

        ok = not (sin_dueno or sin_fila)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {len(montos)} central(es) en [{t_sob}], "
          f"{len(duenos)} en [{t_ce}]")
        L(f"        {con_dueno} con monto y con dueño; {sin_plata} sin monto "
          f"(esas pueden no tener dueño)")
        if sin_dueno:
            L(f"        {len(sin_dueno)} central(es) con MONTO y el dueño "
              f"VACÍO o en 0:")
            for n, v in sorted(sin_dueno, key=lambda x: -abs(x[1]))[:15]:
                L(f"          {n[:34]:<36} {fmt_monto(v)}")
            if len(sin_dueno) > 15:
                L(f"          ... y {len(sin_dueno) - 15} más")
        if sin_fila:
            L(f"        {len(sin_fila)} central(es) con MONTO que NO están en "
              f"[{t_ce}]:")
            for n, v in sorted(sin_fila, key=lambda x: -abs(x[1]))[:15]:
                L(f"          {n[:34]:<36} {fmt_monto(v)}")
            if len(sin_fila) > 15:
                L(f"          ... y {len(sin_fila) - 15} más")
        if not ok:
            L("        Esa plata no tiene a quién pagarse ni a quién cobrarse.")
            L("        Hay que completar el propietario en la planilla de")
            L("        origen y volver a actualizar el Access.")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_centrales=len(montos), n_duenos=len(duenos),
                    con_dueno=con_dueno, sin_plata=sin_plata,
                    sin_dueno=[(n, v) for n, v in
                               sorted(sin_dueno, key=lambda x: -abs(x[1]))[:40]],
                    sin_fila=[(n, v) for n, v in
                              sorted(sin_fila, key=lambda x: -abs(x[1]))[:40]])

    if tipo == "pago_por_empresa":
        # La planilla 1 y la 9 (o la 4) son dos calculos PARALELOS del mismo
        # pago: la 1 lo saca por empresa y concepto directo, la 9 lo saca
        # prorrateando por retiros y despues lo agrupa. Tienen que dar lo
        # mismo, y hasta ahora nadie lo comparaba.
        #
        # Detalles que importan:
        #  - Los nombres de concepto NO se escriben igual: la 1 dice
        #    "CO ERNC" y la 9 "CO_ERNC". Por eso clave_concepto().
        #  - El monto de la 9 corresponde a la columna PAGA de la 1, no a
        #    RECIBE. Verificado en 2409: contra PAGA los 664 pares cuadran
        #    (peor diferencia 58 pesos); contra RECIBE solo 41 de 664.
        #  - La 1 tiene conceptos que la 9 no (los "ID", y CCA/CO/SC_SSCC que
        #    viven en la planilla 4). Solo se comparan los que estan en las
        #    DOS; los que sobran de un lado se informan, no fallan.
        det, res = c["detalle"], c["resumen"]
        r_det = self.rutas.get(det["archivo"])
        r_res = self.rutas.get(res["archivo"])
        if r_det is None or r_res is None:
            L(f"  ? {c['desc']}: falta uno de los archivos")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")

        # --- lado detalle (la planilla 9 o 4: una fila por retiro) ---
        dd = leer_columnas_rapido(
            r_det, det["hoja"],
            [det["col_concepto"], det["col_empresa"], det["col_monto"]],
            det["fila_inicio"], L)
        if dd is None:
            return dict(base, estado="SIN DATOS")
        cc, ce, cm = (x.upper() for x in
                      (det["col_concepto"], det["col_empresa"], det["col_monto"]))
        pdet, n_det = {}, 0
        for f in sorted(dd.get(cc, {})):
            con, emp = dd[cc].get(f), dd.get(ce, {}).get(f)
            if not isinstance(con, str) or not con.strip():
                continue
            if not isinstance(emp, str) or not emp.strip():
                continue
            v = dd.get(cm, {}).get(f)
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                continue
            k = (clave_concepto(con), clave_concepto(emp))
            pdet[k] = pdet.get(k, 0.0) + float(v)
            n_det += 1
        L(f"        {det['nombre']}: {n_det} fila(s) -> {len(pdet)} par(es) "
          f"(concepto, empresa)")

        # --- lado resumen (la planilla 1: ya viene por empresa y concepto) ---
        dr = leer_columnas_rapido(
            r_res, res["hoja"],
            [res["col_concepto"], res["col_empresa"], res["col_monto"]],
            res["fila_inicio"], L)
        if dr is None:
            return dict(base, estado="SIN DATOS")
        rc, re_, rm = (x.upper() for x in
                       (res["col_concepto"], res["col_empresa"], res["col_monto"]))
        pres, n_res = {}, 0
        for f in sorted(dr.get(rc, {})):
            con, emp = dr[rc].get(f), dr.get(re_, {}).get(f)
            if not isinstance(con, str) or not con.strip():
                continue
            if not isinstance(emp, str) or not emp.strip():
                continue
            v = dr.get(rm, {}).get(f)
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                continue
            k = (clave_concepto(con), clave_concepto(emp))
            pres[k] = pres.get(k, 0.0) + float(v)
            n_res += 1
        L(f"        {res['nombre']}: {n_res} fila(s) -> {len(pres)} par(es)")

        # Solo los conceptos que estan en los dos lados.
        con_det = {k[0] for k in pdet}
        con_res = {k[0] for k in pres}
        comunes = con_det & con_res
        solo_det = sorted(con_det - con_res)
        solo_res = sorted(con_res - con_det)
        L(f"        conceptos en común: {len(comunes)}  "
          f"({', '.join(sorted(comunes))})")
        if solo_res:
            L(f"        solo en {res['nombre']}: {', '.join(solo_res)}"
              f"   (no se comparan)")
        if solo_det:
            L(f"        solo en {det['nombre']}: {', '.join(solo_det)}"
              f"   (no se comparan)")

        tol = c.get("tolerancia", TOL_PAGO_EMPRESA)
        # El signo de la columna PAGA de la planilla 1 NO es consistente: es
        # positivo para los conceptos de la planilla 9 (CPF, CSF, CTF, CRA,
        # REA, CO ERNC) y NEGATIVO para los de la planilla 4 (CCA, CO,
        # SC_SSCC). Verificado en 2409, en la misma columna E de la misma hoja.
        # Como la convencion cambia dentro del mismo archivo, no se puede
        # distinguir "convencion" de "error de signo", asi que se compara la
        # MAGNITUD y se informa aparte cuantos pares venian con signo opuesto.
        absoluto = c.get("absoluto", True)
        difs, faltan_res, faltan_det = [], [], []
        n_ok, n_signo = 0, 0
        for k, v in pdet.items():
            if k[0] not in comunes:
                continue
            if k not in pres:
                # Un par que falta con monto ~0 no es un problema: la
                # planilla 1 lista la empresa con 0 y la 9 simplemente no
                # tiene filas para ella en ese concepto.
                if abs(v) > tol:
                    faltan_res.append((k, v))
                continue
            vr = pres[k]
            if absoluto:
                if v * vr < 0:
                    n_signo += 1
                d = abs(v) - abs(vr)
            else:
                d = v - vr
            if abs(d) > tol:
                difs.append((k, v, vr, d))
            else:
                n_ok += 1
        for k, v in pres.items():
            if k[0] in comunes and k not in pdet and abs(v) > tol:
                faltan_det.append((k, v))

        ok = not (difs or faltan_res or faltan_det)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {n_ok} par(es) cuadran dentro de {fmt_monto(tol)}"
          + ("  (se compara la magnitud)" if absoluto else ""))
        if n_signo:
            L(f"        {n_signo} par(es) venían con el signo opuesto, que es "
              f"convención de la planilla 1 y no un error.")
        if difs:
            L(f"        {len(difs)} par(es) con diferencia mayor:")
            for (co, em), vd, vr, d in sorted(difs, key=lambda x: -abs(x[3]))[:15]:
                L(f"          {co:<12} {em[:24]:<26} "
                  f"{det['nombre']} {fmt_monto(vd):>16}  vs  "
                  f"{res['nombre']} {fmt_monto(vr):>16}   dif {fmt_monto(d)}")
            if len(difs) > 15:
                L(f"          ... y {len(difs) - 15} más")
        for (co, em), v in faltan_res[:10]:
            L(f"          {co} / {em[:26]}: está en {det['nombre']} "
              f"({fmt_monto(v)}) y NO en {res['nombre']}")
        if len(faltan_res) > 10:
            L(f"          ... y {len(faltan_res) - 10} más")
        for (co, em), v in faltan_det[:10]:
            L(f"          {co} / {em[:26]}: está en {res['nombre']} "
              f"({fmt_monto(v)}) y NO en {det['nombre']}")
        if len(faltan_det) > 10:
            L(f"          ... y {len(faltan_det) - 10} más")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_ok=n_ok, n_difs=len(difs), n_signo=n_signo,
                    comunes=sorted(comunes),
                    solo_detalle=solo_det, solo_resumen=solo_res,
                    difs=[(f"{co} / {em}", vd, vr, d) for (co, em), vd, vr, d in
                          sorted(difs, key=lambda x: -abs(x[3]))[:40]],
                    faltan_resumen=[f"{co} / {em}" for (co, em), _ in faltan_res[:40]],
                    faltan_detalle=[f"{co} / {em}" for (co, em), _ in faltan_det[:40]])

    if tipo == "bloque_contra_origen":
        # Compara UN bloque del destino (SC o CO) contra su origen filtrado por
        # embalses: cantidad de filas y suma del monto. Es lo unico que dice si
        # los datos pegados son los de este mes; el resto de V10 solo mira que
        # el destino sea coherente consigo mismo.
        # Se hace por bloque separado a proposito, para poder decir CUAL de los
        # dos hay que volver a traer y no actualizar los dos al azar.
        de, org = c["destino"], c["origen"]
        r_dest = self.rutas.get(de["archivo"])
        r_orig = self.rutas.get(org["archivo"])
        if r_dest is None or r_orig is None:
            L(f"  ? {c['desc']}: falta uno de los archivos")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        permitidas = {clave_central(x) for x in CENTRALES_EMBALSE}

        # --- destino: solo las filas de este tipo ---
        dd = leer_columnas_rapido(
            r_dest, de["hoja"],
            [de["col_tipo"], de["col_monto"], de["col_central"]],
            de["fila_inicio"], L)
        if dd is None:
            return dict(base, estado="SIN DATOS")
        objetivo = normalizar(de["tipo"])
        n_dest, suma_dest, cent_dest = 0, 0.0, set()
        for f in sorted(dd.get(de["col_tipo"].upper(), {})):
            if normalizar(dd[de["col_tipo"].upper()][f]) != objetivo:
                continue
            n_dest += 1
            v = dd.get(de["col_monto"].upper(), {}).get(f)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                suma_dest += float(v)
            nom = dd.get(de["col_central"].upper(), {}).get(f)
            if isinstance(nom, str) and nom.strip():
                cent_dest.add(clave_central(nom))

        # --- origen: solo los embalses ---
        do = leer_columnas_rapido(
            r_orig, org["hoja"], [org["col_central"], org["col_monto"]],
            org["fila_inicio"], L)
        if do is None:
            return dict(base, estado="SIN DATOS")
        n_orig, suma_orig, cent_orig = 0, 0.0, set()
        for f in sorted(do.get(org["col_central"].upper(), {})):
            nom = do[org["col_central"].upper()][f]
            if not isinstance(nom, str) or not nom.strip():
                continue
            k = clave_central(nom)
            if k not in permitidas:
                continue
            n_orig += 1
            cent_orig.add(k)
            v = do.get(org["col_monto"].upper(), {}).get(f)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                suma_orig += float(v)

        dif = suma_dest - suma_orig
        ok_filas = n_dest == n_orig
        ok_monto = abs(dif) <= TOLERANCIA
        solo_d = sorted(cent_dest - cent_orig)
        solo_o = sorted(cent_orig - cent_dest)
        ok = ok_filas and ok_monto and not solo_d and not solo_o
        etq = de["tipo"]
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        destino ({etq}): {n_dest} fila(s), "
          f"{fmt_monto(suma_dest)}   [{de['hoja']} col {de['col_monto']}]")
        L(f"        origen  : {n_orig} fila(s), {fmt_monto(suma_orig)}   "
          f"[{Path(r_orig).name}, {org['hoja']} col {org['col_monto']}]")
        if not ok_filas:
            L(f"        FALTAN o SOBRAN filas: {n_dest - n_orig:+d}")
        if not ok_monto:
            L(f"        diferencia de monto: {fmt_monto(dif)}")
        for k in solo_o[:10]:
            L(f"          está en el origen y NO en el destino: {k}")
        for k in solo_d[:10]:
            L(f"          está en el destino y NO en el origen: {k}")
        if not ok:
            L(f"        >>> HAY QUE VOLVER A TRAER LOS {etq} <<<")
            L(f"            (no hace falta tocar el otro bloque)")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    bloque=etq, n_destino=n_dest, n_origen=n_orig,
                    suma_destino=suma_dest, suma_origen=suma_orig,
                    diferencia=dif, solo_destino=solo_d[:40],
                    solo_origen=solo_o[:40])

    if tipo == "suma_fila":
        # Cada fila del bloque de prorrata tiene que sumar el 100%: escrito
        # como 1 o como 100, segun la planilla. Se aceptan los dos, pero se
        # avisa si un mismo archivo mezcla, porque eso ya es un error.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        cols = expandir_columnas(c["rango"])
        f_ini = int(c["fila_inicio"])
        datos = leer_columnas_rapido(ruta, c["hoja"], cols, f_ini, L)
        if datos is None:
            return dict(base, estado="SIN DATOS")
        # Se usa una columna de referencia para saber cuales filas EXISTEN:
        # una fila sin prorrata no es lo mismo que una fila que no existe.
        # Con col_tipo se puede decir si la fila mala es de SC o de CO, que es
        # lo que hace falta para saber cual bloque volver a traer.
        col_tipo = c.get("col_tipo")
        tipos = {}
        if col_tipo:
            dt = leer_columnas_rapido(ruta, c["hoja"], [col_tipo], f_ini, L)
            if dt is not None:
                tipos = {f: (str(v).strip() if isinstance(v, str) else v)
                         for f, v in dt.get(col_tipo.upper(), {}).items()}

        def de_quien(f):
            t = tipos.get(f)
            return str(t) if t else "?"

        col_ref = c.get("col_referencia")
        filas_ref = None
        if col_ref:
            dref = leer_columnas_rapido(ruta, c["hoja"], [col_ref], f_ini, L)
            if dref is not None:
                filas_ref = {f for f, v in dref.get(col_ref.upper(), {}).items()
                             if isinstance(v, str) and v.strip()
                             and clave_central(v) not in ("", "0")}
        if filas_ref is None:
            filas_ref = set()
            for col in cols:
                filas_ref |= set(datos.get(col, {}))
            filas_ref = {f for f in filas_ref if f >= f_ini}

        malas, ceros, escalas = [], [], {}
        for f in sorted(filas_ref):
            vals = [datos.get(col, {}).get(f) for col in cols]
            nums = [float(v) for v in vals
                    if isinstance(v, (int, float)) and not isinstance(v, bool)]
            total = sum(nums)
            # Una fila que suma 0 es VALIDA: no reparte nada. Se cuenta para
            # dejarlo a la vista, pero no hace fallar. Y no entra en la
            # deteccion de escala, porque el 0 no dice si la planilla trabaja
            # en 1 o en 100.
            if not nums or abs(total) <= TOL_PRORRATA:
                ceros.append(f)
                continue
            # ¿1 o 100? se acepta el que quede mas cerca.
            cerca = min(TOTALES_PRORRATA, key=lambda t: abs(total - t))
            tol = TOL_PRORRATA * max(1.0, cerca)
            if abs(total - cerca) <= tol:
                escalas[cerca] = escalas.get(cerca, 0) + 1
            else:
                malas.append((f, total))

        mezcla = len(escalas) > 1
        ok = not malas and not mezcla
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {len(filas_ref)} fila(s) en {c['rango']}{f_ini} hacia abajo")
        for esc, n in sorted(escalas.items()):
            L(f"        {n} fila(s) suman {esc:g}")
        if mezcla:
            L("        OJO: el archivo MEZCLA filas que suman 1 con filas que")
            L("             suman 100. Una de las dos está mal.")
        # Se agrupa por bloque para poder decir cual hay que volver a traer.
        por_bloque = {}
        for f, t in malas:
            por_bloque.setdefault(de_quien(f), []).append((f, t))
        for f, t in malas[:15]:
            L(f"          fila {f} ({de_quien(f)}): suma {t!r}")
        if len(malas) > 15:
            L(f"          ... y {len(malas) - 15} fila(s) más")
        if malas and col_tipo:
            resumen = ", ".join(f"{len(v)} de {k}"
                                for k, v in sorted(por_bloque.items()))
            L(f"        las filas malas son: {resumen}")
            solo = [k for k in por_bloque if k != "?"]
            if len(solo) == 1:
                L(f"        >>> HAY QUE VOLVER A TRAER LOS {solo[0]} <<<")
                L(f"            (no hace falta tocar el otro bloque)")
        if ceros:
            L(f"        {len(ceros)} fila(s) suman 0 (no reparten nada, "
              f"está permitido)")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_filas=len(filas_ref), escalas=dict(escalas),
                    malas=[(f, t, de_quien(f)) for f, t in malas[:40]],
                    ceros=ceros[:40], mezcla=mezcla,
                    por_bloque={k: len(v) for k, v in por_bloque.items()})

    if tipo == "centrales_en_lista":
        # Dos cosas, segun los flags:
        #   exigir  -> toda central de la columna tiene que estar en la lista
        #   sufijo  -> avisar de las "-numero" que NO estan en la lista, que es
        #              la senal de que apareció una unidad de embalse nueva
        lista = CENTRALES_EMBALSE
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        permitidas = {clave_central(x) for x in lista}
        col = c["columna"].upper()
        datos = leer_columnas_rapido(ruta, c["hoja"], [col],
                                     c["fila_inicio"], L)
        if datos is None:
            return dict(base, estado="SIN DATOS")
        vistas, fuera, sospechosas = {}, {}, {}
        for f in sorted(datos.get(col, {})):
            v = datos[col][f]
            if not isinstance(v, str) or not v.strip() or v.startswith("#"):
                continue
            nom = v.strip()
            k = clave_central(nom)
            if k in ("", "0"):
                continue
            if k in permitidas:
                vistas[k] = nom
            else:
                fuera.setdefault(k, (nom, f))
                if RE_UNIDAD_CENTRAL.search(nom):
                    sospechosas.setdefault(k, (nom, f))

        malos = []
        if c.get("exigir"):
            malos = list(fuera.values())
        avisar = list(sospechosas.values()) if c.get("avisar_sufijo") else []
        faltan = [x for x in lista if clave_central(x) not in vistas] \
            if c.get("exigir_todas") else []

        ok = not (malos or avisar or faltan)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        lista: {len(lista)} central(es) de embalse")
        L(f"        en {col}{c['fila_inicio']} hacia abajo: "
          f"{len(vistas)} de la lista, {len(fuera)} fuera de la lista")
        for nom, f in malos[:15]:
            L(f"          {col}{f}: «{nom[:34]}» no es una central de embalse")
        if len(malos) > 15:
            L(f"          ... y {len(malos) - 15} más")
        for nom, f in avisar[:15]:
            L(f"          {col}{f}: «{nom[:34]}» termina en «-número» y NO está")
            L(f"                   en la lista. ¿Es una unidad nueva?")
        if len(avisar) > 15:
            L(f"          ... y {len(avisar) - 15} más")
        for x in faltan[:15]:
            L(f"          falta: {x} no aparece en ninguna fila")
        if len(faltan) > 15:
            L(f"          ... y {len(faltan) - 15} más")
        if ok:
            L("        todas las centrales cuadran con la lista")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_lista=len(lista), n_vistas=len(vistas), n_fuera=len(fuera),
                    fuera=[f"{n} ({col}{fi})" for n, fi in malos[:40]],
                    sufijo=[f"{n} ({col}{fi})" for n, fi in avisar[:40]],
                    faltan=faltan[:40])

    if tipo == "sobrecosto_por_fila":
        # Recalcula el sobrecosto desde sus componentes y lo compara con la
        # columna que lo trae ya calculado:
        #     sobrecosto = (CV - CMg) * Generacion * USD
        # Se compara el TOTAL, que es lo pedido, pero tambien fila por fila:
        # si dos filas se equivocan en sentidos opuestos el total cuadra y el
        # archivo igual esta mal, y sin el detalle por fila no habria como
        # saber DONDE mirar.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        cols = c["columnas"]
        orden = ["cv", "cmg", "gen", "usd", "resultado"]
        letras = [cols[k] for k in orden]
        datos = leer_columnas_rapido(ruta, c["hoja"], letras,
                                     c["fila_inicio"], L)
        if datos is None:
            return dict(base, estado="SIN DATOS")

        def num(letra, fila):
            v = datos.get(letra.upper(), {}).get(fila)
            return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        # Se recorren todas las filas donde haya ALGO, no solo las de la
        # columna del resultado: una fila con componentes y sin resultado es
        # justamente uno de los errores que hay que cazar.
        filas = set()
        for letra in letras:
            filas |= set(datos.get(letra.upper(), {}))
        filas = sorted(f for f in filas if f >= int(c["fila_inicio"]))

        tol_fila = c.get("tolerancia_fila", TOL_SOBRECOSTO_FILA)
        suma_calc = suma_esp = 0.0
        n_ok = 0
        difs, incompletas = [], []
        for f in filas:
            cv, cmg, gen = num(cols["cv"], f), num(cols["cmg"], f), num(cols["gen"], f)
            usd, esp = num(cols["usd"], f), num(cols["resultado"], f)
            if None in (cv, cmg, gen, usd):
                # Sin componentes no se puede recalcular. Solo molesta si la
                # fila SI trae un sobrecosto: ahi hay algo que no cuadra.
                if esp is not None and abs(esp) > TOLERANCIA:
                    faltan = [cols[k] for k, v in
                              (("cv", cv), ("cmg", cmg), ("gen", gen), ("usd", usd))
                              if v is None]
                    incompletas.append((f, esp, "+".join(faltan)))
                continue
            calc = (cv - cmg) * gen * usd
            suma_calc += calc
            if esp is None:
                incompletas.append((f, None, cols["resultado"]))
                continue
            suma_esp += esp
            d = calc - esp
            if abs(d) > tol_fila:
                difs.append((f, calc, esp, d, cv, cmg, gen, usd))
            else:
                n_ok += 1

        dif_total = suma_calc - suma_esp
        ok_total = abs(dif_total) <= max(TOLERANCIA, tol_fila)
        ok = ok_total and not difs and not incompletas
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        ({cols['cv']} − {cols['cmg']}) × {cols['gen']} × "
          f"{cols['usd']}   vs   {cols['resultado']}")
        L(f"        {len(filas)} fila(s) leídas, {n_ok} cuadran fila a fila")
        L(f"        recalculado : {fmt_monto(suma_calc)}")
        L(f"        columna {cols['resultado']}    : {fmt_monto(suma_esp)}")
        L(f"        diferencia  : {fmt_monto(dif_total)}"
          f"   (máximo aceptado {fmt_monto(max(TOLERANCIA, tol_fila))})")
        if difs:
            L(f"        {len(difs)} fila(s) no cuadran (tolerancia por fila "
              f"{fmt_monto(tol_fila)}):")
            for f, calc, esp, d, cv, cmg, gen, usd in sorted(
                    difs, key=lambda x: -abs(x[3]))[:15]:
                L(f"          fila {f:>5}: recalc {fmt_monto(calc):>16} vs "
                  f"{fmt_monto(esp):>16}  dif {fmt_monto(d):>14}")
                L(f"                     {cols['cv']}={cv} {cols['cmg']}={cmg} "
                  f"{cols['gen']}={gen} {cols['usd']}={usd}")
            if len(difs) > 15:
                L(f"          ... y {len(difs) - 15} fila(s) más")
            if ok_total:
                L("        OJO: el TOTAL cuadra pero hay filas que no. Se están")
                L("             compensando entre ellas, así que el archivo está")
                L("             mal aunque la suma dé bien.")
        for f, esp, falta in incompletas[:15]:
            if esp is None:
                L(f"          fila {f}: tiene componentes pero {falta} está vacía")
            else:
                L(f"          fila {f}: trae {fmt_monto(esp)} pero falta {falta}")
        if len(incompletas) > 15:
            L(f"          ... y {len(incompletas) - 15} fila(s) incompleta(s) más")
        if ok:
            L(f"        las {n_ok} filas cuadran una por una")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    suma_calculada=suma_calc, suma_esperada=suma_esp,
                    diferencia=dif_total, n_filas=len(filas), n_ok=n_ok,
                    n_difs=len(difs),
                    difs=[(f, ca, es, d) for f, ca, es, d, *_ in
                          sorted(difs, key=lambda x: -abs(x[3]))[:40]],
                    incompletas=[(f, e, x) for f, e, x in incompletas[:40]])

    if tipo == "matriz_al_dia":
        # La matriz tiene que estar armada con las empresas de AHORA. Si se
        # actualizaron los datos y no se corrio CuadroPago, la matriz queda
        # con las del mes anterior y nadie lo nota: sigue mostrando numeros.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        m = leer_matriz_pago(ruta, c["hoja"], L)
        if m is None or not m["pagan"] or not m["reciben"]:
            L("  >> no se pudo leer la matriz, o esta vacia. "
              "Seguramente falta correr Cuadro de pagos.")
            return dict(base, estado="NO CUADRA", n_pagan=0, n_reciben=0)
        L(f"        matriz: {len(m['pagan'])} que pagan x "
          f"{len(m['reciben'])} que reciben, "
          f"{len(m['montos'])} par(es) con monto")

        # La verdad de quien paga y quien recibe esta en la tabla I:K.
        datos = leer_columnas_rapido(ruta, c["hoja"], ["I", "J", "K"],
                                     c["fila_tabla"], L)
        if datos is None:
            return dict(base, estado="SIN DATOS")
        esp_pagan, esp_reciben = {}, {}
        for f in sorted(datos.get("I", {})):
            nom = datos["I"][f]
            if not isinstance(nom, str) or not nom.strip():
                continue
            k = normalizar(nom)
            if k in ("", "0"):
                continue
            j = datos.get("J", {}).get(f)
            kk = datos.get("K", {}).get(f)
            if isinstance(j, (int, float)) and not isinstance(j, bool):
                esp_pagan[k] = nom.strip()
            if isinstance(kk, (int, float)) and not isinstance(kk, bool):
                esp_reciben[k] = nom.strip()
        L(f"        tabla I:K: {len(esp_pagan)} con monto en J (pagan), "
          f"{len(esp_reciben)} con monto en K (reciben)")

        hay_pagan = {normalizar(x) for x in m["pagan"]}
        hay_reciben = {normalizar(x) for x in m["reciben"]}
        faltan_p = [esp_pagan[k] for k in esp_pagan if k not in hay_pagan]
        sobran_p = [x for x in m["pagan"] if normalizar(x) not in esp_pagan]
        faltan_r = [esp_reciben[k] for k in esp_reciben if k not in hay_reciben]
        sobran_r = [x for x in m["reciben"] if normalizar(x) not in esp_reciben]

        # Y el nombre definido tiene que cubrir justo esa matriz, si no la
        # consulta lee de mas o de menos (falto apretar Actualiza Rango).
        nombre = c.get("nombre_definido")
        aviso_rango, rango_ok = None, True
        if nombre:
            crudo = leer_nombre_definido(ruta, nombre)
            fin = celdas_de_rango(crudo)
            # El rango excluye la fila y columna de totales: por eso -1 no,
            # sino que termina justo en la ultima empresa.
            if fin is None:
                aviso_rango = f"no se pudo leer el nombre '{nombre}': {crudo!r}"
                rango_ok = False
            elif (fin[0], fin[1]) != (m["ultima_fila"], m["ultima_col"]):
                aviso_rango = (
                    f"{nombre} termina en fila {fin[0]} columna {fin[1]}, "
                    f"y la matriz termina en fila {m['ultima_fila']} columna "
                    f"{m['ultima_col']}. Falta apretar «Actualiza Rango».")
                rango_ok = False
            else:
                aviso_rango = (f"{nombre} cubre justo la matriz "
                               f"(hasta fila {fin[0]}, columna {fin[1]})")

        ok = not (faltan_p or sobran_p or faltan_r or sobran_r) and rango_ok
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        for et, lista in (("PAGAN, falta en la matriz", faltan_p),
                          ("PAGAN, sobra en la matriz", sobran_p),
                          ("RECIBEN, falta en la matriz", faltan_r),
                          ("RECIBEN, sobra en la matriz", sobran_r)):
            for n in sorted(lista)[:10]:
                L(f"          {et}: {n[:40]}")
            if len(lista) > 10:
                L(f"          ... y {len(lista) - 10} mas ({et})")
        if aviso_rango:
            L(f"        {aviso_rango}")
        if not ok and (faltan_p or sobran_p or faltan_r or sobran_r):
            L("        La matriz no corresponde a los datos de ahora: "
              "falta correr «Cuadro de pagos».")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_pagan=len(m["pagan"]), n_reciben=len(m["reciben"]),
                    faltan_pagan=faltan_p[:40], sobran_pagan=sobran_p[:40],
                    faltan_reciben=faltan_r[:40], sobran_reciben=sobran_r[:40],
                    aviso_rango=aviso_rango)

    if tipo == "cprt_al_dia":
        # El CPRT sale de la dinamica. Si no se refresco, sus pares son los
        # del mes anterior aunque la matriz ya este bien.
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")
        m = leer_matriz_pago(ruta, c["hoja_matriz"], L)
        if m is None or not m["montos"]:
            L("  >> no se pudo leer la matriz; sin ella no hay con que "
              "comparar el CPRT.")
            return dict(base, estado="SIN DATOS")
        d = leer_columnas_rapido(ruta, c["hoja_cprt"],
                                 ["B", "E", "G"], c["fila_cprt"], L)
        if d is None:
            return dict(base, estado="SIN DATOS")
        pares_cprt = {}
        for f in sorted(d.get("B", {})):
            deu, acr = d["B"].get(f), d.get("E", {}).get(f)
            if not (isinstance(deu, str) and deu.strip()):
                continue
            if not (isinstance(acr, str) and acr.strip()):
                continue
            monto = d.get("G", {}).get(f)
            pares_cprt[(normalizar(deu), normalizar(acr))] = (
                deu.strip(), acr.strip(),
                float(monto) if isinstance(monto, (int, float)) else None)
        L(f"        CPRT: {len(pares_cprt)} par(es)   |   "
          f"matriz: {len(m['montos'])} par(es) con monto")

        # Direccion 1: todo par del CPRT tiene que existir en la matriz, y
        # con el mismo monto redondeado. Esto caza los pares viejos que
        # quedaron de un refresco anterior.
        fantasmas, difs = [], []
        for k, (deu, acr, monto) in pares_cprt.items():
            if k not in m["montos"]:
                fantasmas.append(f"{deu} -> {acr}")
                continue
            if monto is not None and abs(round(m["montos"][k]) - round(monto)) > 1:
                difs.append((deu, acr, monto, m["montos"][k]))
        # Direccion 2: solo se exigen los pares CLARAMENTE grandes. El CPRT
        # descarta los montos chicos, y el corte exacto no esta documentado
        # (en 2312 los excluidos llegaban a 8,7 y el menor incluido era 22,3).
        # Pedir los grandes caza un refresco viejo sin inventar el umbral.
        perdidos = [f"{p} -> {r}" for (p, r), v in m["montos"].items()
                    if abs(v) >= UMBRAL_PAR_SEGURO and (p, r) not in pares_cprt]
        ok = not (fantasmas or difs or perdidos)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        for n in sorted(fantasmas)[:10]:
            L(f"          en el CPRT y NO en la matriz: {n[:52]}")
        if len(fantasmas) > 10:
            L(f"          ... y {len(fantasmas) - 10} mas")
        for n in sorted(perdidos)[:10]:
            L(f"          en la matriz y NO en el CPRT: {n[:52]}")
        if len(perdidos) > 10:
            L(f"          ... y {len(perdidos) - 10} mas")
        for deu, acr, mc, mm in sorted(difs, key=lambda x: -abs(x[2] - x[3]))[:10]:
            L(f"          {deu[:18]} -> {acr[:18]}: CPRT {fmt_monto(mc)} "
              f"vs matriz {fmt_monto(mm)}")
        if len(difs) > 10:
            L(f"          ... y {len(difs) - 10} diferencia(s) mas")
        if not ok:
            L("        El CPRT no corresponde a la matriz de ahora: "
              "falta refrescar la tabla dinámica.")
        else:
            L(f"        los {len(pares_cprt)} pares del CPRT calzan con la matriz")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    n_cprt=len(pares_cprt), n_matriz=len(m["montos"]),
                    fantasmas=fantasmas[:40], perdidos=perdidos[:40],
                    difs=[(a, b, x, y) for a, b, x, y in difs[:40]])

    if tipo == "mismas_empresas":
        # Los dos lados tienen que traer EXACTAMENTE las mismas empresas: ni
        # una de mas ni una de menos. Los 0 y los vacios se descartan porque
        # son lo que arrastra una formula que sobra, y eso esta permitido.
        # Se compara el conjunto, no el largo ni la ultima: asi da igual el
        # orden y da igual que una columna tenga cola de ceros.
        la, lb = c["lado_a"], c["lado_b"]
        ra = self.rutas.get(la["archivo"])
        rb = self.rutas.get(lb["archivo"])
        if ra is None or rb is None:
            L(f"  ? {c['desc']}: falta uno de los archivos")
            return dict(base, estado="SIN DATOS")
        L(f"  .. {c['desc']}")

        def leer_lado(ruta, esp):
            d = leer_columnas_rapido(ruta, esp["hoja"], [esp["col"]],
                                     esp.get("fila_inicio", 1), L)
            if d is None:
                return None
            col = esp["col"].upper()
            out, dups, basura = {}, [], 0
            for f in sorted(d.get(col, {})):
                v = d[col][f]
                if isinstance(v, str) and v.startswith("#"):
                    basura += 1          # error de formula
                    continue
                if not isinstance(v, str) or not v.strip():
                    basura += 1          # numero (0) o vacio
                    continue
                k = normalizar(v)
                if k in ("", "0"):
                    basura += 1
                    continue
                if k in out:
                    dups.append((v.strip(), f))
                    continue
                out[k] = (v.strip(), f)
            nombre = esp.get("nombre", f"{esp['hoja']} {col}")
            L(f"        {nombre}: {len(out)} empresa(s)"
              + (f", {basura} fila(s) con 0/vacío descartadas" if basura else ""))
            return {"emp": out, "dups": dups, "nombre": nombre}

        a = leer_lado(ra, la)
        b = leer_lado(rb, lb)
        if a is None or b is None:
            return dict(base, estado="SIN DATOS")

        solo_a = [a["emp"][k][0] for k in a["emp"] if k not in b["emp"]]
        solo_b = [b["emp"][k][0] for k in b["emp"] if k not in a["emp"]]
        dups = []
        if c.get("sin_duplicados", True):
            dups = ([(a["nombre"], n, f) for n, f in a["dups"]]
                    + [(b["nombre"], n, f) for n, f in b["dups"]])
        ok = not (solo_a or solo_b or dups)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        if solo_a:
            L(f"        {len(solo_a)} empresa(s) en {a['nombre']} que FALTAN "
              f"en {b['nombre']}:")
            for n in sorted(solo_a)[:15]:
                L(f"          {n[:44]}")
            if len(solo_a) > 15:
                L(f"          ... y {len(solo_a) - 15} más")
        if solo_b:
            L(f"        {len(solo_b)} empresa(s) en {b['nombre']} que SOBRAN "
              f"(no están en {a['nombre']}):")
            for n in sorted(solo_b)[:15]:
                L(f"          {n[:44]}")
            if len(solo_b) > 15:
                L(f"          ... y {len(solo_b) - 15} más")
        for et, n, f in dups[:10]:
            L(f"        REPETIDA en {et}: {n[:36]} (fila {f})")
        if len(dups) > 10:
            L(f"        ... y {len(dups) - 10} repetición(es) más")
        if ok:
            L(f"        las {len(a['emp'])} empresas son las mismas en los dos lados")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    solo_a=solo_a[:40], solo_b=solo_b[:40],
                    duplicadas=[(et, n, f) for et, n, f in dups[:40]],
                    n_a=len(a["emp"]), n_b=len(b["emp"]),
                    nombre_a=a["nombre"], nombre_b=b["nombre"])

    if tipo == "pertenencia":
        ruta = self.rutas.get(c["archivo"])
        if ruta is None:
            L(f"  ? {c['desc']}: falta el archivo")
            return dict(base, estado="SIN DATOS")
        datos = leer_columnas_rapido(
            ruta, c["hoja"], list(c["origen"]) + list(c["destino"]),
            c.get("fila_inicio", 1), L)
        if datos is None:
            return dict(base, estado="SIN DATOS")

        def textos(cols, recolectar_dups=False):
            """{clave: (nombre, col, fila)}. Descarta vacios, errores y los
            "0" que aparecen cuando sobran formulas."""
            out, dups = {}, []
            for col in cols:
                for f in sorted(datos.get(col, {})):
                    v = datos[col][f]
                    if not isinstance(v, str) or not v.strip() or v.startswith("#"):
                        continue
                    k = normalizar(v)
                    if k in ("", "0"):
                        continue
                    if k in out:
                        if recolectar_dups:
                            dups.append((v.strip(), col, f))
                        continue
                    out[k] = (v.strip(), col, f)
            return (out, dups) if recolectar_dups else out

        orig = textos(c["origen"])
        dest, dups = textos(c["destino"], recolectar_dups=True)
        if not c.get("sin_duplicados_destino"):
            dups = []
        faltan = [orig[k] for k in orig if k not in dest]
        ok = not (faltan or dups)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {len(orig)} empresa(s) distintas en {'+'.join(c['origen'])}, "
          f"{len(dest)} en {'+'.join(c['destino'])}")
        for nombre, col, f in faltan[:20]:
            L(f"          falta: {nombre[:34]:<36} (está en {col}{f})")
        if len(faltan) > 20:
            L(f"          ... y {len(faltan) - 20} más")
        for nombre, col, f in dups[:20]:
            L(f"          REPETIDA en {col}{f}: {nombre[:34]}")
        if len(dups) > 20:
            L(f"          ... y {len(dups) - 20} repetición(es) más")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    faltan=[n for n, _, _ in faltan[:40]],
                    duplicadas=[f"{n} ({col}{f})" for n, col, f in dups[:40]],
                    n_origen=len(orig), n_destino=len(dest))

    if tipo == "ultimo_igual":
        ra = self.rutas.get(c["archivo_a"])
        rb = self.rutas.get(c["archivo_b"])
        if ra is None or rb is None:
            L(f"  ? {c['desc']}: falta uno de los archivos")
            return dict(base, estado="SIN DATOS")
        da = leer_columnas_rapido(ra, c["hoja_a"], [c["col_a"]],
                                  c.get("fila_a", 1), L)
        db = (da if (ra == rb and c["hoja_a"] == c["hoja_b"])
              else leer_columnas_rapido(rb, c["hoja_b"], [c["col_b"]],
                                        c.get("fila_b", 1), L))
        if da is None or db is None:
            return dict(base, estado="SIN DATOS")
        if ra == rb and c["hoja_a"] == c["hoja_b"] and c["col_b"] not in da:
            da2 = leer_columnas_rapido(ra, c["hoja_a"], [c["col_a"], c["col_b"]],
                                       c.get("fila_a", 1), L)
            da = db = da2 if da2 is not None else da
        fa, va = ultimo_significativo(da, c["col_a"])
        fb, vb = ultimo_significativo(db, c["col_b"])
        if va is None or vb is None:
            L(f"  ? {c['desc']}: una de las columnas no tiene datos útiles "
              f"({c['col_a']}={va}, {c['col_b']}={vb})")
            return dict(base, estado="SIN DATOS")
        if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
            ok = abs(va - vb) <= TOLERANCIA
        else:
            ok = normalizar(va) == normalizar(vb)
        L(f"  {'OK ' if ok else '>> '}{c['desc']}")
        L(f"        {c['col_a']}{fa} = {str(va)[:38]!r}")
        L(f"        {c['col_b']}{fb} = {str(vb)[:38]!r}")
        return dict(base, estado="OK" if ok else "NO CUADRA",
                    a=f"{c['col_a']}{fa}", b=f"{c['col_b']}{fb}",
                    valor_a=va, valor_b=vb)

    L(f"  ? comprobación de tipo desconocido: {tipo}")
    return dict(base, estado="SIN DATOS")
