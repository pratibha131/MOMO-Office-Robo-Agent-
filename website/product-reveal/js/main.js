(() => {
  const canvas = document.getElementById('scene');
  const wrapper = document.getElementById('assembly-track');

  const scene = new THREE.Scene();
  scene.fog = new THREE.Fog(0x001538, 6, 16);

  const camera = new THREE.PerspectiveCamera(42, window.innerWidth / window.innerHeight, 0.1, 100);
  camera.position.set(0, 0.2, 8);

  const renderer = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: true });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.setSize(window.innerWidth, window.innerHeight);

  // ---------- Lighting ----------
  scene.add(new THREE.AmbientLight(0x9fb8dd, 0.55));

  const key = new THREE.DirectionalLight(0xffffff, 1.1);
  key.position.set(3, 5, 4);
  scene.add(key);

  const rim = new THREE.PointLight(0x6fd1ff, 1.6, 20);
  rim.position.set(-4, 1, -3);
  scene.add(rim);

  const fill = new THREE.PointLight(0x0b5ed7, 0.9, 20);
  fill.position.set(2, -3, 3);
  scene.add(fill);

  // ---------- Product group ----------
  const group = new THREE.Group();
  scene.add(group);

  const navyMat  = new THREE.MeshStandardMaterial({ color: 0x00265e, metalness: 0.6, roughness: 0.35 });
  const blueMat  = new THREE.MeshStandardMaterial({ color: 0x0b5ed7, metalness: 0.45, roughness: 0.3 });
  const cyanMat  = new THREE.MeshStandardMaterial({ color: 0x6fd1ff, metalness: 0.3, roughness: 0.25, emissive: 0x1a5c86, emissiveIntensity: 0.25 });
  const whiteMat = new THREE.MeshStandardMaterial({ color: 0xf3f7fc, metalness: 0.2, roughness: 0.4 });
  const darkMat  = new THREE.MeshStandardMaterial({ color: 0x0a1220, metalness: 0.7, roughness: 0.2 });

  function makePart(geometry, material, finalPos, finalRotY = 0) {
    const mesh = new THREE.Mesh(geometry, material);
    mesh.position.copy(finalPos);
    group.add(mesh);

    // scattered starting point: random direction, generous distance, tumbled rotation
    const dir = new THREE.Vector3(
      (Math.random() * 2 - 1),
      (Math.random() * 2 - 1),
      (Math.random() * 2 - 1)
    ).normalize();
    const distance = 4.5 + Math.random() * 3.5;
    const startPos = finalPos.clone().add(dir.multiplyScalar(distance));

    const startQuat = new THREE.Quaternion().setFromEuler(
      new THREE.Euler(Math.random() * Math.PI * 2, Math.random() * Math.PI * 2, Math.random() * Math.PI * 2)
    );
    const endQuat = new THREE.Quaternion().setFromEuler(new THREE.Euler(0, finalRotY, 0));

    return { mesh, startPos, endPos: finalPos.clone(), startQuat, endQuat };
  }

  const parts = [
    makePart(new THREE.CylinderGeometry(1.15, 1.25, 0.4, 40), navyMat, new THREE.Vector3(0, -1.55, 0)),
    makePart(new THREE.CylinderGeometry(0.9, 0.9, 2.2, 40), blueMat, new THREE.Vector3(0, -0.15, 0)),
    makePart(new THREE.TorusGeometry(0.95, 0.09, 16, 48), cyanMat, new THREE.Vector3(0, 0.95, 0), Math.PI / 2),
    makePart(new THREE.SphereGeometry(0.92, 40, 24, 0, Math.PI * 2, 0, Math.PI / 1.9), whiteMat, new THREE.Vector3(0, 1.55, 0)),
    makePart(new THREE.BoxGeometry(0.55, 0.35, 0.06), darkMat, new THREE.Vector3(0, 0.15, 0.92)),
    makePart(new THREE.CylinderGeometry(0.09, 0.09, 0.06, 24), cyanMat, new THREE.Vector3(0, -0.55, 0.92)),
  ];
  parts[2].mesh.rotation.x = Math.PI / 2; // torus lies flat as a collar by default

  // fix torus base orientation so its end rotation matches "lying flat"
  parts[2].endQuat.setFromEuler(new THREE.Euler(Math.PI / 2, 0, 0));

  // ---------- Scroll-driven progress ----------
  function easeOutCubic(t) { return 1 - Math.pow(1 - t, 3); }

  function getProgress() {
    const rect = wrapper.getBoundingClientRect();
    const total = wrapper.offsetHeight - window.innerHeight;
    if (total <= 0) return 1;
    const scrolled = -rect.top;
    return Math.min(1, Math.max(0, scrolled / total));
  }

  const heroCopy = document.querySelector('.hero-copy');
  const scrollCue = document.querySelector('.scroll-cue');

  let clock = new THREE.Clock();

  function animate() {
    requestAnimationFrame(animate);
    const delta = clock.getDelta();

    const raw = getProgress();
    const t = easeOutCubic(raw);

    parts.forEach((p) => {
      p.mesh.position.lerpVectors(p.startPos, p.endPos, t);
      THREE.Quaternion.slerp(p.startQuat, p.endQuat, p.mesh.quaternion, t);
    });

    // gentle continuous tumble that settles as assembly completes
    group.rotation.y += delta * (0.25 - 0.18 * t);
    group.rotation.x = Math.sin(performance.now() * 0.00015) * 0.05 * (1 - t * 0.6);

    // camera moves in slightly as the product completes
    camera.position.z = 8 - 2 * t;
    camera.position.y = 0.2 - 0.1 * t;
    camera.lookAt(0, 0, 0);

    // fade hero copy out and scroll cue out as user scrolls into the build
    const fade = 1 - Math.min(1, raw / 0.35);
    heroCopy.style.opacity = fade;
    heroCopy.style.transform = `translateY(calc(-50% + ${(1 - fade) * 30}px))`;
    scrollCue.style.opacity = raw < 0.05 ? 1 : 0;

    renderer.render(scene, camera);
  }
  animate();

  // ---------- Resize ----------
  function onResize() {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(window.innerWidth, window.innerHeight);
  }
  window.addEventListener('resize', onResize);
  window.addEventListener('orientationchange', onResize);
})();
