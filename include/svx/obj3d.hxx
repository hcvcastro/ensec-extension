/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This file incorporates work covered by the following license notice:
 *
 *   Licensed to the Apache Software Foundation (ASF) under one or more
 *   contributor license agreements. See the NOTICE file distributed
 *   with this work for additional information regarding copyright
 *   ownership. The ASF licenses this file to you under the Apache
 *   License, Version 2.0 (the "License"); you may not use this file
 *   except in compliance with the License. You may obtain a copy of
 *   the License at http://www.apache.org/licenses/LICENSE-2.0 .
 */

#ifndef INCLUDED_SVX_OBJ3D_HXX
#define INCLUDED_SVX_OBJ3D_HXX

#include <svx/svdoattr.hxx>
#include <svx/svdobj.hxx>
#include <svx/def3d.hxx>
#include <svx/svdpage.hxx>
#include <svx/deflt3d.hxx>
#include <vcl/bitmap.hxx>
#include <svx/svx3ditems.hxx>
#include <svx/xflclit.hxx>
#include <svl/itemset.hxx>
#include <basegfx/range/b3drange.hxx>
#include <basegfx/polygon/b3dpolypolygon.hxx>
#include <basegfx/point/b2dpoint.hxx>
#include <basegfx/polygon/b2dpolygon.hxx>
#include <basegfx/vector/b2enums.hxx>
#include <svx/svxdllapi.h>


// Forward declaration


class SfxPoolItem;
class Viewport3D;
class E3dScene;
class E3dPolyScene;

namespace basegfx { class B3DPolyPolygon; }
namespace sdr { namespace properties {
    class BaseProperties;
    class E3dProperties;
    class E3dCompoundProperties;
    class E3dExtrudeProperties;
    class E3dLatheProperties;
    class E3dSphereProperties;
}}

/*************************************************************************
|*
|* GeoData relevant for undo actions
|*
\************************************************************************/

class E3DObjGeoData : public SdrObjGeoData
{
public:
    basegfx::B3DRange           maLocalBoundVol;    // surrounding volume of the object
    basegfx::B3DHomMatrix       maTransformation;   // lokal transformations

    E3DObjGeoData() {}
};

/*************************************************************************
|*
|* List for 3D objects
|*
\************************************************************************/

class E3dObjList : public SdrObjList
{
    E3dObjList &operator=(const E3dObjList& rSrcList) SAL_DELETED_FUNCTION;

public:
    TYPEINFO_OVERRIDE();
    E3dObjList(SdrModel* pNewModel = 0, SdrPage* pNewPage = 0, E3dObjList* pNewUpList = 0);
    SVX_DLLPUBLIC virtual ~E3dObjList();

    virtual E3dObjList* Clone() const SAL_OVERRIDE;

    virtual void NbcInsertObject(SdrObject* pObj, size_t nPos=SAL_MAX_SIZE, const SdrInsertReason* pReason=NULL) SAL_OVERRIDE;
    virtual void InsertObject(SdrObject* pObj, size_t nPos=SAL_MAX_SIZE, const SdrInsertReason* pReason=NULL) SAL_OVERRIDE;
    virtual SdrObject* NbcRemoveObject(size_t nObjNum) SAL_OVERRIDE;
    virtual SdrObject* RemoveObject(size_t nObjNum) SAL_OVERRIDE;

protected:
    SVX_DLLPUBLIC E3dObjList(const E3dObjList& rSrcList);
};

/*************************************************************************
|*
|* Base class for 3D objects
|*
\************************************************************************/

class SVX_DLLPUBLIC E3dObject : public SdrAttrObj
{
private:
    // to allow sdr::properties::E3dProperties access to StructureChanged()
    friend class sdr::properties::E3dProperties;

    // Allow everything for E3dObjList and E3dDragMethod
    friend class E3dObjList;
    friend class E3dDragMethod;

 protected:
    virtual sdr::properties::BaseProperties* CreateObjectSpecificProperties() SAL_OVERRIDE;

    E3dObjList                  maSubList;          // child objects

    basegfx::B3DRange           maLocalBoundVol;    // surrounding volume of the object (from the geometry generation)
    basegfx::B3DHomMatrix       maTransformation;   // local transformation
    basegfx::B3DHomMatrix       maFullTransform;    // global transformation (including. parents)

    // Flags
    bool            mbTfHasChanged          : 1;
    bool            mbIsSelected            : 1;

 public:
    void SetBoundVolInvalid();

 protected:
    void SetTransformChanged();
    virtual void NewObjectInserted(const E3dObject* p3DObj);
    virtual void StructureChanged();
    basegfx::B3DRange RecalcBoundVolume() const;

protected:
    // E3dObject is only a helper class (for E3DScene and E3DCompoundObject)
    // and no instances should be created from anyone, so i move the constructors
    // to protected area
    E3dObject();
    E3dObject(bool bIsFromChart);

public:
    TYPEINFO_OVERRIDE();
    virtual void RecalcSnapRect() SAL_OVERRIDE;
    virtual void SetRectsDirty(bool bNotMyself = false) SAL_OVERRIDE;

    virtual ~E3dObject();

    virtual sal_uInt32  GetObjInventor() const SAL_OVERRIDE;
    virtual sal_uInt16  GetObjIdentifier() const SAL_OVERRIDE;

    virtual void    TakeObjInfo(SdrObjTransformInfoRec& rInfo) const SAL_OVERRIDE;

    virtual void        NbcSetLayer(SdrLayerID nLayer) SAL_OVERRIDE;

    virtual void        SetObjList(SdrObjList* pNewObjList) SAL_OVERRIDE;
    virtual void        SetPage(SdrPage* pNewPage) SAL_OVERRIDE;
    virtual void        SetModel(SdrModel* pNewModel) SAL_OVERRIDE;
    virtual void        NbcMove(const Size& rSize) SAL_OVERRIDE;
    virtual void NbcResize(const Point& rRef, const Fraction& xFact, const Fraction& yFact) SAL_OVERRIDE;
    virtual SdrObjList* GetSubList() const SAL_OVERRIDE;

    // Insert 3D object into the group; transfer to other owner!
    void Insert3DObj(E3dObject* p3DObj);
    void Remove3DObj(E3dObject* p3DObj);

    E3dObject* GetParentObj() const;
    virtual E3dScene* GetScene() const;

    const basegfx::B3DRange& GetBoundVolume() const;
    void InvalidateBoundVolume();

    // calculate complete transformation including all parents
    const basegfx::B3DHomMatrix& GetFullTransform() const;

    // get and (re)set transformation matrix
    const basegfx::B3DHomMatrix& GetTransform() const { return maTransformation;}
    virtual void NbcSetTransform(const basegfx::B3DHomMatrix& rMatrix);
    virtual void SetTransform(const basegfx::B3DHomMatrix& rMatrix);

    // 2D rotations, are implemented as a rotation around the Z axis
    // which is vertical to the screen, plus a shift of the scene.
    // This means that also the scene (E3dScene) must define this
    // routine as virtual in its class.
    virtual void NbcRotate(const Point& rRef, long nAngle, double sn, double cs) SAL_OVERRIDE;

    // get wireframe polygon for local object. No transform is applied.
    basegfx::B3DPolyPolygon CreateWireframe() const;

    // TakeObjName...() is for the display in the UI, for example "3 frames selected".
    virtual OUString TakeObjNameSingul() const SAL_OVERRIDE;
    virtual OUString TakeObjNamePlural() const SAL_OVERRIDE;
    virtual E3dObject* Clone() const SAL_OVERRIDE;
    E3dObject& operator=( const E3dObject& rObj );

    virtual SdrObjGeoData *NewGeoData() const SAL_OVERRIDE;
    virtual void          SaveGeoData(SdrObjGeoData& rGeo) const SAL_OVERRIDE;
    virtual void          RestGeoData(const SdrObjGeoData& rGeo) SAL_OVERRIDE;

    // get/set the selection
    bool GetSelected() const { return mbIsSelected; }
    void SetSelected(bool bNew);

    // break up
    virtual bool IsBreakObjPossible();
    virtual SdrAttrObj* GetBreakObj();
};

/*************************************************************************
|*
|* Class for all compound objects (Cube, Lathe, Scene, Extrude)
|* This class saves some ISA queries and accelerates the behaviour
|* significantly, because all the attributes etc. are kept through this.
|* The polygons may only keep attributes if they are directly
|* subordinated to the scene.
|*
\************************************************************************/

class SVX_DLLPUBLIC E3dCompoundObject : public E3dObject
{
private:
    // to allow sdr::properties::E3dCompoundProperties access to SetGeometryValid()
    friend class sdr::properties::E3dCompoundProperties;
    friend class sdr::properties::E3dExtrudeProperties;
    friend class sdr::properties::E3dLatheProperties;
    friend class sdr::properties::E3dSphereProperties;

protected:
    virtual sdr::properties::BaseProperties* CreateObjectSpecificProperties() SAL_OVERRIDE;

    // material of the object
    Color                   aMaterialAmbientColor;

    // attributes for geometry creation
    bool                    bCreateNormals              : 1;
    bool                    bCreateTexture              : 1;

protected:
    void SetDefaultAttributes(E3dDefaultAttributes& rDefault);

    // convert given basegfx::B3DPolyPolygon to screen coor
    basegfx::B2DPolyPolygon TransformToScreenCoor(const basegfx::B3DPolyPolygon& rCandidate);

public :
    TYPEINFO_OVERRIDE();

    E3dCompoundObject();
    E3dCompoundObject(E3dDefaultAttributes& rDefault);
    virtual ~E3dCompoundObject();

    virtual basegfx::B2DPolyPolygon TakeXorPoly() const SAL_OVERRIDE;
    virtual sal_uInt32 GetHdlCount() const SAL_OVERRIDE;
    virtual void    AddToHdlList(SdrHdlList& rHdlList) const SAL_OVERRIDE;

    // DoubleSided: sal_True/FALSE
    bool GetDoubleSided() const
        { return static_cast<const SfxBoolItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_DOUBLE_SIDED)).GetValue(); }

    // NormalsKind: 0 == sal_False/sal_False, 1 == sal_True/sal_False, else == sal_True/TRUE
    sal_uInt16 GetNormalsKind() const
        { return static_cast<const Svx3DNormalsKindItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_NORMALS_KIND)).GetValue(); }

    // NormalsInvert: sal_True/FALSE
    bool GetNormalsInvert() const
        { return static_cast<const SfxBoolItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_NORMALS_INVERT)).GetValue(); }

    // TextureProjX: 0 == sal_False/sal_False, 1 == sal_True/sal_False, else == sal_True/TRUE
    sal_uInt16 GetTextureProjectionX() const
        { return static_cast<const Svx3DTextureProjectionXItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_TEXTURE_PROJ_X)).GetValue(); }

    // TextureProjY: 0 == sal_False/sal_False, 1 == sal_True/sal_False, else == sal_True/TRUE
    sal_uInt16 GetTextureProjectionY() const
        { return static_cast<const Svx3DTextureProjectionYItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_TEXTURE_PROJ_Y)).GetValue(); }

    // Shadow3D: sal_True/FALSE
    bool GetShadow3D() const
        { return static_cast<const SfxBoolItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_SHADOW_3D)).GetValue(); }

    // MaterialColor: Color
    Color GetMaterialColor() const
        { return static_cast<const SvxColorItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_MAT_COLOR)).GetValue(); }

    // MaterialEmission: Color
    Color GetMaterialEmission() const
        { return static_cast<const SvxColorItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_MAT_EMISSION)).GetValue(); }

    // MaterialSpecular: Color
    Color GetMaterialSpecular() const
        { return static_cast<const SvxColorItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_MAT_SPECULAR)).GetValue(); }

    // MaterialSpecularIntensity:
    sal_uInt16 GetMaterialSpecularIntensity() const
        { return static_cast<const SfxUInt16Item&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_MAT_SPECULAR_INTENSITY)).GetValue(); }

    // TextureFilter: sal_True/FALSE
    bool GetTextureFilter() const
        { return static_cast<const SfxBoolItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_TEXTURE_FILTER)).GetValue(); }

    // #i28528#
    // Added extra Item (Bool) for chart2 to be able to show reduced line geometry
    bool GetReducedLineGeometry() const
        { return static_cast<const Svx3DReducedLineGeometryItem&>(GetObjectItemSet().Get(SDRATTR_3DOBJ_REDUCED_LINE_GEOMETRY)).GetValue(); }

    virtual sal_uInt16 GetObjIdentifier() const SAL_OVERRIDE;
    virtual void RecalcSnapRect() SAL_OVERRIDE;

    // set/get parameters for geometry creation
    bool GetCreateNormals() const { return bCreateNormals; }

    bool GetCreateTexture() const { return bCreateTexture; }

    virtual E3dCompoundObject* Clone() const SAL_OVERRIDE;

    // material of the object
    const Color& GetMaterialAmbientColor() const { return aMaterialAmbientColor; }

    bool IsAOrdNumRemapCandidate(E3dScene*& prScene) const;
};

#endif // INCLUDED_SVX_OBJ3D_HXX

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
